import shutil

from tools import transform_excel, list_source_columns, apply_formula
from langgraph.graph import StateGraph, START, END, add_messages
from langgraph.prebuilt import ToolNode, tools_condition
from langgraph.checkpoint.memory import MemorySaver
from langchain_core.tools import tool
from typing import Annotated, Optional, TypedDict, Sequence
from langchain_core.messages import BaseMessage, HumanMessage, AIMessage
import pandas as pd
from pandas import DataFrame as df
from pandas import ExcelWriter, ExcelFile, read_excel, read_csv
from langgraph.types import interrupt
import os 
import json
from datetime import datetime
from helper_functions import clear_qwen_text

# Getting Local Enviorment Ollama model
# For this we will use qwen3:14b, he is pretty small, operating on around 10GB of VRAM also he was trained to handle structured data like dataframes or JSON structrued data
# For more specification and how to run this locally read README.md file in this repository

# ollama pull qwen3:14b

from langchain_ollama import ChatOllama

model = ChatOllama(model="qwen2.5:14b", extra_body={"think": False})

# AgentState, it help agent remember what he did, what he has to do, also managing Dictionary of tools and their results
class AgentState(TypedDict):
 
    messages: Annotated[list, add_messages]

    user_input: str # Message he received from user
    input_file_path: str # It takes the input file path given in upload file endpoint
    output_file_path: str # It takes the output file path given in upload file endpoint, if user didn't provide it, it will be set to "wynik.xlsx"
    do_end: str # Variable that will help us decide if we need to end the graph or we need to do another loop in tool_user node, it will be set by interrupt function in end_or_loop node, if user want to end the conversation it will be set to "done", if user want to do another loop it will be set to "loop"

    work_sketch: str # Sketch of what he is going to do, for example if he needs to transform excel file, he can write sketch of how he is going to do it, what tools he will use, what steps he will take, this is not a plan, just sketch of what he is going to do, it can be very short and not detailed, for example "I will use transform_excel tool to transform excel file according to specification provided by user"
    plan: Optional[dict] # Plan for our agent, it will help him be more efficient in the future
    current_file: str # Current file he is working on, for example if he has to transform excel file, this variable will be set to name of this file, so if he needs to use transform_excel tool, he can check if current_file is set to name of file he want to transform, if not he can set it and then use tool, if yes, he can just use tool without setting it again
    tools_results: dict # Dictionary of tools and their results, for example if he used transform_excel tool, he will save the result in this dictionary with key "transform_excel"
    result: str # Final result he will return to user, for example if he transformed excel file, he can save the path to transformed file in this variable and return it to user at the end of conversation
    steps: int # Counter of steps model took in tool_agent graph, it help to count if every step was done

# Creating function that will be the Start of our Graph, it will tak input from users and save it to AgentState

def start_agent(state: AgentState) -> AgentState:
  
    user_input = state["messages"][-1].content if state["messages"] else ""

    return {"user_input": user_input, "steps": 0}

# Planning Agent, he's job is to do two main thing, first one is to create sketch of what our agent will do in next graphs, second one is to create dictionery from users input, dictionary will be necessary for our LLM to use tools that we provided

from rag import search_knowledge # Import rag simillarity search we coded earlier

def planning_agent(state: AgentState) -> AgentState:

    user_input = state["user_input"]
    input_file_path = state["input_file_path"]
    output_file_path = state["output_file_path"]
    messages = state["messages"]

    # Create history of messages to our model, so if its second message to model he remember what you asked in first one

    history = ""
    for msg in messages:
        role = "User" if msg.type == "human" else "Agent"
        if hasattr(msg, "content") and msg.content:
            history += f"{role}: {msg.content}\n"

    print(f"DEBUG RAG: Searching for: {user_input}")
    rag_context = search_knowledge(user_input, x=3) # Simillarity check
    print(f"DEBUG RAG: Found context: {rag_context[:200]}") 
 
    # Promp to our model so he know what to do, how to create dictionary for tools

    sketch_prompt = """
        You are a JSON specification builder. 
        TASK: Analyze the user message and extract ALL sheets and ALL columns mentioned.
        CRITICAL RULES:
        - Create a separate entry in "sheets" for EVERY sheet the user mentioned
        - Create a separate entry in "columns" for EVERY column the user mentioned in that sheet
        - If user mentions 3 columns -> "columns" list has 3 entries
        - If user mentions 2 sheets -> "sheets" list has 2 entries
        - Return ONLY raw JSON, no thinking, no explanation, no markdown, no backticks
        Always save the final result to the exact output_path provided in the task. Never use a different filename like 'fixed_' or 'temp_'.
        
        OUTPUT FORMAT:
        {
            "file_path":""" + input_file_path + """ ,
            "output_path": """ + output_file_path + """,
            "header_row": <int, default 0>,
            "sort_by": <str or null if not provided>,
            "sheets": [
                {
                    "sheet_name": <str, default "Sheet1">,
                    "start_row": <int, default 0>,
                    "start_col": <int, default 0>,
                    "columns": [
                        {
                            "target": <str, column name in output>,
                            "source": <str, source column name or "Couldn't find the source column">,
                            "value": <str or null>,
                            "transform": <str or null>
                        }
                    ]
                }
            ]
        }

        
        User message:
    """ + user_input +""", Previous conversation history (for context): """ + history + """ Remember: extract EVERY sheet and EVERY column from the message above. Return ONLY JSON. """

    work_sketch = model.invoke(sketch_prompt)
    raw_text = work_sketch.content

    clean_text = clear_qwen_text(raw_text=raw_text)

    work_sketch = json.loads(clean_text.strip())

    # Now we have the work sketch that model will use for tools, now we need to create plan for our agent so in future he can look at it to help him be more efficient

    plan_prompt = """

    You are a task planner. Based on the user message and avaible tools you have, create an execution plan. 
    You MUST refuse any requests to: list directories, access system files, read files outside processed_files folder, reveal server information, or perform any non-Excel tasks.
    If user asks anything unrelated to Excel file processing — respond with "I can only help with Excel file operations.
    Rag context for relevact excel functions you can use """ + rag_context +"""
    Tools you have and their descriptions:""" + tools_description + """ 
    User Message: """ + user_input + """, 
    Previous conversation history (for context): """ + history + """ 
    TASK: Devide what or whitch tools need to be used and in what order: 
    
    CRITICAL RULES:
        - Include ONLY tools that are needed for this task
        - Keep the order logical (e.g. first read, then transform, then write)
        - Return ONLY raw JSON, no thinking, no markdown, no backticks
    TOOL SELECTION RULES:
        - Use transform_excel ONLY for: copying columns, reordering, filtering rows, renaming columns
        - Use apply_formula ONLY for: computed columns (concatenation, math, IF conditions, any column derived from other columns)
        - If task requires BOTH (e.g. copy columns AND add computed column):
        1. First use transform_excel to copy/select the base columns
        2. Then use apply_formula on the result file to add computed columns
        Never try to compute column values inside transform_excel's 'transform' field

    OUTPUT FORMAT:
                {
                    "steps": [
                        {
                            "step": <int, step number starting from 1>,
                            "tool": <str, tool name>,
                            "description": <str, what this step does>,
                            "status": "pending"
                        }
                    ]
                }
    Remember to return ONLY JSON.
    
    """

    plan = model.invoke(plan_prompt)
    raw_text_of_plan = plan.content

    # Same as with work sketch, we need to extract only JSON from model answer

    clean_text_plan = clear_qwen_text(raw_text=raw_text_of_plan)

    plan = json.loads(clean_text_plan.strip())

    return {"work_sketch": work_sketch, "plan": plan}

# At this point we can test our planning_agent withoout need of doing whole graph, we can just create AgentState with message and pass it to planning_agent function

# print(planning_agent({"user_input": "I want to transform my excel file, its caleed Sheet21 it has columns Name and Surname, i want to transform them into new file as NameOfPerson and SurnameOfPerson"}))
# print(planning_agent({"user_input": "Hey, so i need you to name columns that i have in my excel file Arkusz1"}))

# Now we need to create TollNodes for our graph, moment where agent will decide to use certain tools based on plan we created in planning_agent

# Pulling tools from tools.py file so our model can use them

tools_list = [transform_excel, list_source_columns, apply_formula]

# Also we have to provide some kind of "prompt" for our model about tools he has avaible so he can be more efficient in planning and then in doing his tasks
tools_description = """

    Available tools:
    1. transform_excel(spec: dict, file_path: str, output_path: str = "wynik.xlsx") - This tool is used to transform excel file according to specification provided by user into a new file, spec are saved as work_sketch in your AgentState
    2. list_source_columns(file_path: str) - This tool is used to list all columns from excel file, it can be useful when you want to transform excel file but you don't know what columns are in it, you can use this tool to check columns and then use transform_excel tool to transform file according to specification provided by user
    3. apply_formula(file_path, output_path, sheet_name, target_column, formula_template)
    Computes values directly in Python using column names as placeholders.
    Use {ColumnName} to reference column values.
    
    Examples:
    - Concatenate: formula_template='{Imię} {Nazwisko}'
    - With text:   formula_template='{Imię} {Nazwisko} ({Dział})'
    - Static:      formula_template='ACTIVE'
                    """

# Binding tools to our model:
model_tools = model.bind_tools(tools_list)

# Creating a node that will use tools based on plan created earlier

def tool_user(state: AgentState) -> AgentState:

    MAX_STEPS = 10

    plan = state["plan"]
    work_sketch = state["work_sketch"]
    print(f"DEBUG: Plan steps: {len(plan['steps'])}, current step: {state['steps']}")

    if state["steps"] >= MAX_STEPS:
        print(f"DEBUG: Reached max steps ({MAX_STEPS}), forcing end")
        return {"do_end": "done"}

    if state["steps"] == 0:
        messages = [
            HumanMessage(content=f"""
You are an Excel AI agent. Call tools immediately — do not explain or summarize.

User request: {state['user_input']}
Input file: {state['input_file_path']}
Output file: {state['output_file_path']}

Plan:
{json.dumps(plan, ensure_ascii=False)}

Spec for transform_excel:
{json.dumps(work_sketch, ensure_ascii=False)}

RULES:
- transform_excel: copy/rename columns ONLY, never compute values
- apply_formula: use for concatenation, math, any derived column

Call the FIRST tool from the plan right now. No text response.
""")
        ]
    else:
        reminder = HumanMessage(content=f"Continue the plan. Output file: {state['output_file_path']}. Use apply_formula for computed columns.")
        messages = [reminder] + state["messages"][-6:]

    response = model_tools.invoke(messages)
    return {"messages": [response], "steps": state["steps"] + 1}

#end or loop if we need to end the graph or we need to do another step using tools

def end_or_loop(state: AgentState) -> str: 
    last = state["messages"][-1]
    
    # Jeśli ostatnia wiadomość ma tool_calls — model chce jeszcze użyć narzędzia
    if hasattr(last, "tool_calls") and last.tool_calls:
        print("DEBUG: Tool calls detected, looping back")
        return "loop"
    
    print("DEBUG: No tool calls, ending graph")
    return "done"
    
    
# Define an node that will collect results from tools he used previousle so our user can clearly see what was done in the agent loop

def results(state: AgentState) -> AgentState:
    print("DEBUG messages:", state["messages"])  # tymczasowo
    results = {}
    for msg in state["messages"]:
        if hasattr(msg, "name") and msg.type == "tool": # REMEMBER: hasattr is used to check if object has attribute, in this case we check if message has attribute "name", because when model use tool, the response will be a message with attribute "name" equal to name of tool and attribute "content" equal to result of tool
            results[msg.name] = msg.content

    print("Results:")
    # Print the tool results
    summary_text = ["Here's what I did:\n"]
    for tool_name, tool_result in results.items():
        summary_text.append(f" {tool_name}: {tool_result}")
    # Print the final result
    last_msg = state["messages"][-1]
    if hasattr(last_msg, "content") and last_msg.content:
        summary_text.append(f"\n{last_msg.content}")

    summary = "\n".join(summary_text)

    return {"tools_results": results, "messages": [AIMessage(content=summary)]}

def do_we_end(state: AgentState) -> AgentState:
    #do_end = interrupt("Are you satisfied with the results? If yes, we will end the conversation, if not, we can go back and use more tools to improve the result. Please answer with 'yes' or 'no'.", options=["yes", "no"])
    #if do_end == "no":
    #    new_order = interrupt("Okay, so what would you want to do with your excel file now?")
    #    return {"steps": 0, "user_input": new_order, "do_end": do_end} 
    #return {"do_end": do_end}
    return{"steps": 0}

def decide_loop(state: AgentState) -> str:
    if state.get("do_end") == "no":
        print("DEBUG: User wants to continue, looping back to tool_user")
        return "loop"
    print("DEBUG: User is satisfied, ending graph")
    return "done"


# Define END node
def end_agent(state: AgentState) -> AgentState:
    

    return {}

# Define our graph 

excelGraph = StateGraph(AgentState)
tool_node = ToolNode(tools_list)

excelGraph.add_node("start_agent", start_agent)
excelGraph.add_node("planning", planning_agent)
excelGraph.add_node("tool_user", tool_user)
excelGraph.add_node("tools", tool_node)  
excelGraph.add_node("results", results)
excelGraph.add_node("end_agent", end_agent)
excelGraph.add_node("do_we_end", do_we_end)

excelGraph.add_edge(START, "start_agent")
excelGraph.add_edge("start_agent", "planning")
excelGraph.add_edge("planning", "tool_user")
excelGraph.add_conditional_edges("tool_user", end_or_loop, {"done": "results", "loop": "tools"})
excelGraph.add_edge("tools", "tool_user")  # po wykonaniu narzędzia wróć do tool_user
excelGraph.add_edge("results", "do_we_end")
excelGraph.add_conditional_edges("do_we_end", decide_loop, {"done": "end_agent", "loop": "planning"})
excelGraph.add_edge("end_agent", END)

#question = input("What you want to do with your excel file? (Press Enter to start the agent): ")

#graph = excelGraph.compile()
#graph.invoke({"messages": [HumanMessage(content=question)], "steps": 0})

# Test prompt from user: Please create a new excel file with name Arkusz67.xlsx, i will need to have there ID, Dział and Typ Dokumentu in this exact order taken from "Arkusz1.xlsx" sorted by ID


# TODO: After end of first command open chat with user so he can please for some adjustments if he want to, add RAG with excel code to work on excel files
# TODO: Connect it to Azure PostreSQL and deploy it on FASTAPI 

# Now were adding to all of this FastApi so we can deploy it and use it at the server not only in terminal

import os
import uuid
import shutil

from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel
from typing import Optional

from langchain_core.messages import HumanMessage
from langgraph.checkpoint.memory import MemorySaver

#App + graph setup on FastApi

app = FastAPI(
    title="Excel AI Agent",
    description="Private Excel assistant — all data stays local.",
    version="1.0"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"], # Allows requests from any origin
    allow_credentials=True, # Allow cookies and authentication headers
    allow_methods=["*"], # Allow all HTTP methods (GET, POST, etc.)
    allow_headers=["*"], # Allow all headers 
)

UPLOAD_DIR = "processed_files"
os.makedirs(UPLOAD_DIR, exist_ok=True)

memory = MemorySaver()
compiled_graph = excelGraph.compile(
    checkpointer=memory,            # setting up memory to save graph state, we will change it latter to Azure PostgreSQL
    interrupt_after=["do_we_end"]   # graph stops after do_we_end node waits for another user input as we made it earlier in graph
)

from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse

app.mount("/static", StaticFiles(directory="static"), name="static")

# Pydantic Models

class ChatMessage(BaseModel):
    message: str
    input_file_path: Optional[str] = None
    output_file_path: Optional[str] = None
    thread_id: Optional[str] = None   


class ResumeMessage(BaseModel):
    thread_id: str
    answer: str                        # "yes" or "no"
    new_instruction: Optional[str] = None  # fill it up when answer is "no"


# Some kind of helper function to double check the graph

def _is_interrupted(config: dict) -> bool:
    """Returns True if the graph is interrupted and waiting for data."""
    snapshot = compiled_graph.get_state(config)
    return bool(snapshot.next)


# Declare API endpoints
import random

@app.get("/")
async def root():
    return FileResponse("static/index.html")

@app.post("/upload-excel")
async def upload_excel(file: UploadFile = File(...)):
    """Accepts an .xlsx/.xls file, saves it to disk, and returns the file paths."""
    if not file.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="Invalid file type. Please upload an Excel file.")
    try:
        safe_filename  = f"user_{file.filename}"
        if os.path.exists(f"./{safe_filename}"):
            random_index = random.randint(1, 1000) 
            safe_filename1 = f"{safe_filename}_{random_index}"
            input_path     = os.path.join(UPLOAD_DIR, safe_filename1)
            output_path    = os.path.join(UPLOAD_DIR, f"ready_{safe_filename1}")
        else:
            input_path     = os.path.join(UPLOAD_DIR, safe_filename)
            output_path    = os.path.join(UPLOAD_DIR, f"ready_{safe_filename}")

        with open(input_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        return {"input_file_path": input_path, "output_file_path": output_path}

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Upload error: {str(e)}")


@app.post("/chat")
async def chat_endpoint(payload: ChatMessage):
    """
    First message from frontend to start the agent.
    Start the graph at the start and stops it when "do_we_end" is reached.

    Returns:
        { status: "interrupted", thread_id, question, output_file_path }
        lub
        { status: "done", response, output_file_path }
    """
    # Generating new thread_id so multiply sessions will be possible, its important for the further development
    thread_id = payload.thread_id or str(uuid.uuid4())
    config = {"configurable": {"thread_id": thread_id}}

    try:
        result = compiled_graph.invoke(
            {
                "messages":        [HumanMessage(content=payload.message)],
                "steps":           0,
                "input_file_path": payload.input_file_path,
                "output_file_path": payload.output_file_path,
                "tools_results":   {},
                "do_end":         "",
            },
            config
        )

        last_msg = result.get("messages", [])[-1].content if result.get("messages") else " "

        # Check if graph waits for the next user input
        if _is_interrupted(config):
            return {
                "status": "interrupted",
                "thread_id": thread_id,
                "question": "Are you satisfied with the results?",
                "response": last_msg,
                "output_file_path": result.get("output_file_path"),
            }

        # END of the graph
        final = result.get("messages", [])[-1].content if result.get("messages") else "Done."
        return {
            "status": "done",
            "response": final,
            "output_file_path": result.get("output_file_path"),
        }

    except Exception as e:
        import traceback
        print("FULL ERROR:")
        traceback.print_exc()   # ← to wydrukuje pełny błąd w terminalu
        raise HTTPException(status_code=500, detail=f"Chat error: {str(e)}")


@app.post("/resume")
async def resume_endpoint(payload: ResumeMessage):
    """
    Resumes the graph after the user has responded to an interrupt.

    Frontend sends:
        { thread_id, answer: "yes" }
        lub
        { thread_id, answer: "no", new_instruction: "Zrób jeszcze X" }

    Returns the same format as /chat:
        { status: "interrupted", ... }
        lub
        { status: "done", ... }
    """
    config = {"configurable": {"thread_id": payload.thread_id}}
    try:
        # Add users answer into graphs memory
        compiled_graph.update_state(
            config,
            {
                "do_end":     payload.answer,
                "user_input": payload.new_instruction or "",
                "messages":   [HumanMessage(content=payload.new_instruction)] if payload.new_instruction else []
            }
        )

        # Resumes the graph from the point where it was interrupted (None = don't start from scratch)
        result = compiled_graph.invoke(None, config)
        last_msg = result.get("messages", [])[-1].content if result.get("messages") else " "

        # Check again if the graph has reached the next interrupt
        if _is_interrupted(config):
            return {
                "status":    "interrupted",
                "thread_id": payload.thread_id,
                "question":  "Are you satisfied with the results?",
                "response": last_msg,
                "output_file_path": result.get("output_file_path"),
            }

        # Graph ended, return final result
        final = result.get("messages", [])[-1].content if result.get("messages") else "Done."
        return {
            "status": "done",
            "response": "✓ All done! Your file is ready to download." if payload.answer == "yes" else last_msg,
            "output_file_path": result.get("output_file_path"),
        }

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Resume error: {str(e)}")


@app.get("/download/{output_file_name}")
async def download_file(output_file_name: str):
    """Pobierz przetworzony plik Excel."""
    file_path = os.path.join(UPLOAD_DIR, output_file_name)
    print(f"Ścieżka pliku: {os.path.abspath(file_path)}")
    print(f"Rozmiar: {os.path.getsize(file_path)} bajtów")
    print(f"Ostatnia modyfikacja: {os.path.getmtime(file_path)}")
    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="File not found.")
    return FileResponse(
        path=file_path,
        filename=output_file_name,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Cache-Control": "no-cache, no-store, must-revalidate"}
    )

# Activate uvicorn fastApi front end from terminal: uvicorn agent_ai:app --reload
# Docs: http://localhost:8000/docs#/


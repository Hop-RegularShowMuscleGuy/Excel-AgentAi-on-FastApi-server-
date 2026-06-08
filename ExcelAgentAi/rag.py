import os
from langchain_chroma import Chroma
from langchain_ollama import OllamaEmbeddings
from langchain_core.documents import Document

# Build Rag System + temporary code for creating a vector database from excel_functions.md

CHROMA_DIR = "chroma_db" # Later we will change that setup to Azure
KNOWLEDGE_DIR = "local_rag"

# Embending Model
embeddings = OllamaEmbeddings(model="nomic-embed-text") # Ollama doesnt have dedicated embending motel fro qwen we use in main code so we will try to use this one, #ollama pull nomic-embed-text

def get_vectorstore() -> Chroma:
    """Load existing ChromaDB OR create a new one"""
    return Chroma(
        persist_directory=CHROMA_DIR,
        embedding_function=embeddings
    )

def search_knowledge(query: str, x: int = 3):
    """
    Search ChromaDB for relevant Excel function documentation.
    Returns top x results as a single string for use in prompts.
    """
    vectorstore = get_vectorstore()
    results = vectorstore.similarity_search(query, k=x)
    if not results:
        return "No relevant documentation found."
    return "\n---\n".join([doc.page_content for doc in results])

# Make a vector store from prepared .md files in "local_rag"
def build_knowledge_base():
    """
        Read all .txt/.md files from local_rag/ folder and index them into ChromaDB.
        Run this when you add new documentation.
    """

    docs = []
    for filename in os.listdir(KNOWLEDGE_DIR):
        if filename.endswith((".txt", ".md")):
            path = os.path.join(KNOWLEDGE_DIR, filename)
            with open(path, encoding="utf-8") as f:
                content = f.read()
            docs.append(Document(page_content=content, metadata={"source": filename}))

    vectorstore= Chroma(
        persist_directory=CHROMA_DIR,
        embedding_function=embeddings
    )

    vectorstore.add_documents(docs)
    print(f"Succesfully indexed {len(docs)} documents to ChromaDC")


# Acive code below only once than command it

if __name__ == "__main__":
    build_knowledge_base()
import os
import glob
from langchain_community.document_loaders import PyMuPDFLoader
from langchain_text_splitters import RecursiveCharacterTextSplitter
from langchain_chroma import Chroma
from langchain_community.embeddings import HuggingFaceEmbeddings

# Define paths
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DB_DIR = os.path.join(BASE_DIR, "chroma_db")

def find_pdf_files(base_path):
    """Finds all PDF files in the base directory and its subdirectories."""
    search_pattern = os.path.join(base_path, "**", "*.pdf")
    return glob.glob(search_pattern, recursive=True)

def ingest_documents():
    print("Starting document ingestion process...")
    
    # 1. Find all PDFs
    pdf_files = find_pdf_files(BASE_DIR)
    print(f"Found {len(pdf_files)} PDF files.")
    
    if not pdf_files:
        print("No PDF files found. Please ensure PDFs are placed in the directory structure.")
        return

    documents = []
    
    # 2. Extract text from PDFs
    for pdf_path in pdf_files:
        print(f"Loading {pdf_path}...")
        try:
            loader = PyMuPDFLoader(pdf_path)
            docs = loader.load()
            
            # Add metadata about source
            for doc in docs:
                doc.metadata["source_file"] = os.path.basename(pdf_path)
                
            documents.extend(docs)
        except Exception as e:
            print(f"Error loading {pdf_path}: {e}")

    print(f"Total pages extracted: {len(documents)}")

    # 3. Chunk the content
    print("Chunking documents...")
    text_splitter = RecursiveCharacterTextSplitter(
        chunk_size=1000,
        chunk_overlap=200,
        separators=["\n\n", "\n", ".", " ", ""]
    )
    chunks = text_splitter.split_documents(documents)
    print(f"Created {len(chunks)} chunks.")

    # 4. Create Embeddings and Vector Store
    print("Initializing embedding model (all-MiniLM-L6-v2)...")
    embeddings = HuggingFaceEmbeddings(model_name="all-MiniLM-L6-v2")
    
    print("Building Vector Database with Chroma...")
    # Using persist_directory to save locally
    vectorstore = Chroma.from_documents(
        documents=chunks,
        embedding=embeddings,
        persist_directory=DB_DIR
    )
    
    print(f"Successfully ingested data and saved vector store at {DB_DIR}")

if __name__ == "__main__":
    ingest_documents()

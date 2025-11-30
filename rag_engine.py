import os
from dotenv import load_dotenv
from langchain_community.document_loaders import PyMuPDFLoader
from langchain_text_splitters import RecursiveCharacterTextSplitter
from langchain_huggingface import HuggingFaceEmbeddings
from langchain_chroma import Chroma
from langchain_openai import ChatOpenAI
from langchain_classic.chains import create_retrieval_chain
from langchain_classic.chains.combine_documents import create_stuff_documents_chain
from langchain_core.prompts import ChatPrompt
Template

load_dotenv()

# Configuration
PERSIST_DIRECTORY = "./chroma_db"
EMBEDDING_MODEL_NAME = "all-MiniLM-L6-v2"

def load_and_process_pdf(pdf_path):
    """
    Loads a PDF and splits it into chunks.
    """
    loader = PyMuPDFLoader(pdf_path)
    documents = loader.load()
    
    text_splitter = RecursiveCharacterTextSplitter(
        chunk_size=1000,
        chunk_overlap=200,
        add_start_index=True
    )
    chunks = text_splitter.split_documents(documents)
    return chunks

def get_vectorstore(chunks=None):
    """
    Returns the Chroma vector store. If chunks are provided, it adds them to the store.
    """
    embeddings = HuggingFaceEmbeddings(model_name=EMBEDDING_MODEL_NAME)
    
    if chunks:
        vectorstore = Chroma.from_documents(
            documents=chunks,
            embedding=embeddings,
            persist_directory=PERSIST_DIRECTORY
        )
    else:
        vectorstore = Chroma(
            persist_directory=PERSIST_DIRECTORY,
            embedding_function=embeddings
        )
    return vectorstore

def get_rag_chain(vectorstore):
    """
    Creates the RAG chain using OpenRouter/OpenAI interface.
    """
    # OpenRouter Configuration
    llm = ChatOpenAI(
        openai_api_key=os.getenv("OPENROUTER_API_KEY"),
        openai_api_base="https://openrouter.ai/api/v1",
        model_name="openai/gpt-oss-20b:free", # Using a good default, can be changed
        temperature=0
    )

    retriever = vectorstore.as_retriever(search_type="similarity", search_kwargs={"k": 5})

    system_prompt = (
        "You are an assistant for question-answering tasks. "
        "Use the following pieces of retrieved context to answer "
        "the question. If you don't know the answer, say that you "
        "don't know. Use three sentences maximum and keep the "
        "answer concise."
        "\n\n"
        "{context}"
    )

    prompt = ChatPromptTemplate.from_messages(
        [
            ("system", system_prompt),
            ("human", "{input}"),
        ]
    )

    question_answer_chain = create_stuff_documents_chain(llm, prompt)
    rag_chain = create_retrieval_chain(retriever, question_answer_chain)

    return rag_chain

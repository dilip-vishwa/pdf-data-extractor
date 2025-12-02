import os
from dotenv import load_dotenv
from langchain_community.document_loaders import PyMuPDFLoader
from langchain_text_splitters import RecursiveCharacterTextSplitter
from langchain_huggingface import HuggingFaceEmbeddings
from langchain_chroma import Chroma
from langchain_openai import ChatOpenAI
from langchain_classic.chains import create_retrieval_chain
from langchain_classic.chains.combine_documents import create_stuff_documents_chain
from langchain_core.prompts.chat import ChatPromptTemplate

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
        add_start_index=True,
        separators=["[PAGE_BREAK]", "\n\n", "\n", " ", ""]
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
        model_name="openai/gpt-oss-20b:free",
        # model_name="x-ai/grok-4.1-fast:free",
        temperature=0
    )

    retriever = vectorstore.as_retriever(search_type="similarity", search_kwargs={"k": 10})

    system_prompt = (
        "You are an expert about electrical design. Give me answer in short and simple words."
        "No need to give any explanation. Don't give any extra information."
        "You are filling Specification summary sheet of a electrical design project."
        "If answer is expected from given options, then give answer from given options only." 
        "This data is needed to be filled in excel sheet."
        # "If you don't know the answer, then say 'NA'."
        "Try to match with the options if given in prompt"
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

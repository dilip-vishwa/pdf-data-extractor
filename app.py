import logging
logging.basicConfig(level=logging.INFO)
from dotenv.main import load_dotenv
import streamlit as st
import os
import tempfile
from rag_engine import load_and_process_pdf, get_vectorstore, get_rag_chain
load_dotenv()
os.environ["OPENROUTER_API_KEY"] = os.getenv("OPENROUTER_API_KEY")
st.set_page_config(page_title="PDF Data Extractor", layout="wide")

st.title("PDF Data Extractor")

with st.sidebar:
    st.header("Configuration")
    st.markdown("---")
    st.markdown("### Instructions")
    st.markdown("1. Upload a PDF file. Multiple file can be selected for upload in upload dialog box")
    st.markdown("2. Ask questions about the content.")


uploaded_files = st.file_uploader("Upload a PDF", type="pdf", accept_multiple_files=True)

if len(uploaded_files) > 0:
    tmp_paths = []
    for uploaded_file in uploaded_files:
        # Save uploaded file to a temporary file
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            tmp_paths.append(tmp_file.name)

        st.success(f"Uploaded: {uploaded_file.name}")

    
    if "vectorstore" not in st.session_state:
        with st.spinner("Processing PDF... This may take a while for large files."):
            try:
                chunks = []
                for tmp_path in tmp_paths:
                    chunks.extend(load_and_process_pdf(tmp_path))

                st.session_state.vectorstore = get_vectorstore(chunks)
                st.success("PDF Processed and Vector Store Created!")
            except Exception as e:
                st.error(f"Error processing PDF: {e}")
            finally:
                # Cleanup temp file
                for tmp_path in tmp_paths:
                    if os.path.exists(tmp_path):
                        os.remove(tmp_path)

    if "messages" not in st.session_state:
        st.session_state.messages = []

    # Display chat messages
    for message in st.session_state.messages:
        with st.chat_message(message["role"]):
            st.markdown(message["content"])

    # Chat input
    if prompt := st.chat_input("Ask a question about the PDF..."):
        if not os.environ.get("OPENROUTER_API_KEY"):
            st.error("Please enter your OpenRouter API Key in the sidebar.")
        elif "vectorstore" not in st.session_state:
            st.error("Please process the PDF first.")
        else:
            st.session_state.messages.append({"role": "user", "content": prompt})
            with st.chat_message("user"):
                st.markdown(prompt)

            with st.chat_message("assistant"):
                with st.spinner("Thinking..."):
                    try:
                        rag_chain = get_rag_chain(st.session_state.vectorstore)
                        response = rag_chain.invoke({"input": prompt})
                        answer = response["answer"]
                        st.markdown(answer)
                        st.session_state.messages.append({"role": "assistant", "content": answer})
                        
                        # Show sources
                        with st.expander("View Sources"):
                            for i, doc in enumerate(response["context"]):
                                st.markdown(f"**Source {i+1}** (Page {doc.metadata.get('page', 'N/A')}):")
                                st.markdown(f"> {doc.page_content[:200]}...")
                    except Exception as e:
                        st.error(f"Error generating answer: {e}")

else:
    st.info("Please upload a PDF to get started.")

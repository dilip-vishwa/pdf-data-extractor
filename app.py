from excel_utils.excel_save_answer import fill_excel
from excel_utils.excel_save_answer import get_answer_from_llm_response
from excel_utils.excel_prompt import get_prompt_from_excel
import logging
logging.basicConfig(level=logging.INFO)
from dotenv.main import load_dotenv
import streamlit as st
import os
import time
import tempfile
from rag_utils.rag_engine import load_and_process_pdf, get_vectorstore, get_rag_chain
load_dotenv()
os.environ["OPENROUTER_API_KEY"] = os.getenv("OPENROUTER_API_KEY")
st.set_page_config(page_title="PDF Data Extractor", layout="wide")

st.title("PDF Data Extract and put it in given Excel")

with st.sidebar:
    st.header("Configuration")
    st.markdown("---")
    st.markdown("### Instructions")
    st.markdown("1. Upload a PDF file. Multiple file can be selected for upload in upload dialog box")
    st.markdown("2. Upload a Excel file. Wherever data is needed in excel file, mark that cell with green fill color")
    st.markdown("3. Generate Data in excel file")


uploaded_files = st.file_uploader("Upload a PDF", type="pdf", accept_multiple_files=True)

if len(uploaded_files) > 0:
    tmp_paths = []
    for uploaded_file in uploaded_files:
        # Save uploaded pdf file to a temporary file
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            tmp_paths.append(tmp_file.name)

        st.success(f"Uploaded: {uploaded_file.name}")

    
    if "vectorstore" not in st.session_state:
        with st.spinner("Processing PDF... This may take a while for large files."):
            try:
                try:
                    os.rmdir("./chroma_db")
                except:
                    pass
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

    uploaded_excel_file = st.file_uploader("Upload a Excel File", type="xlsx")
    if uploaded_excel_file:
        excel_tmp_paths = ""
        # Save uploaded excel file to a temporary file
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_file:
            tmp_file.write(uploaded_excel_file.getvalue())
            excel_tmp_paths = tmp_file.name

        st.success(f"Uploaded: {uploaded_excel_file.name}")

    if st.button("Generate Excel file with data from PDF"):
        st.write("Generate button clicked!")
        prompts = get_prompt_from_excel(excel_tmp_paths)

        response_from_llm = []
        for prompt in prompts:
            user_prompt = "\n".join(prompt)
            st.session_state.messages.append(user_prompt)
            with st.chat_message("user"):
                st.markdown(prompt)

            with st.chat_message("assistant"):
                with st.spinner("Thinking..."):
                    try:
                        rag_chain = get_rag_chain(st.session_state.vectorstore)
                        response = rag_chain.invoke({"input": user_prompt})
                        response_from_llm.append(response)
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

            time.sleep(3) # for rate limiting

        answer_from_llm_with_excel_coordinates = get_answer_from_llm_response(response_from_llm)
        fill_excel(excel_tmp_paths, answer_from_llm_with_excel_coordinates, "output.xlsx")
        st.info("Excel file generated successfully!")
else:
    st.info("Please upload a PDF to get started.")

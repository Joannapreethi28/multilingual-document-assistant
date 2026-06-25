import streamlit as st
import numpy as np
import faiss
from pypdf import PdfReader
from docx import Document
from sentence_transformers import SentenceTransformer
from io import BytesIO


# -------------------------------
# PAGE CONFIG
# -------------------------------
st.set_page_config(
    page_title="Multilingual RAG Assistant",
    page_icon="🌐",
    layout="wide"
)


# -------------------------------
# LOAD MULTILINGUAL EMBEDDING MODEL
# -------------------------------
@st.cache_resource
def load_embedding_model():
    model = SentenceTransformer("sentence-transformers/paraphrase-multilingual-MiniLM-L12-v2")
    return model


model = load_embedding_model()


# -------------------------------
# TEXT EXTRACTION FUNCTIONS
# -------------------------------
def extract_text_from_pdf(file):
    text = ""
    pdf_reader = PdfReader(file)

    for page in pdf_reader.pages:
        page_text = page.extract_text()
        if page_text:
            text += page_text + "\n"

    return text


def extract_text_from_docx(file):
    text = ""
    document = Document(file)

    for paragraph in document.paragraphs:
        text += paragraph.text + "\n"

    return text


def extract_text_from_txt(file):
    text = file.read().decode("utf-8", errors="ignore")
    return text


def extract_text(uploaded_file):
    file_name = uploaded_file.name.lower()

    if file_name.endswith(".pdf"):
        return extract_text_from_pdf(uploaded_file)

    elif file_name.endswith(".docx"):
        return extract_text_from_docx(uploaded_file)

    elif file_name.endswith(".txt"):
        return extract_text_from_txt(uploaded_file)

    else:
        return ""


# -------------------------------
# CHUNKING FUNCTION
# -------------------------------
def chunk_text(text, chunk_size=500, overlap=100):
    words = text.split()
    chunks = []

    start = 0

    while start < len(words):
        end = start + chunk_size
        chunk = " ".join(words[start:end])

        if chunk.strip():
            chunks.append(chunk)

        start = end - overlap

    return chunks


# -------------------------------
# CREATE FAISS INDEX
# -------------------------------
def create_faiss_index(chunks):
    embeddings = model.encode(chunks, convert_to_numpy=True)

    embeddings = np.array(embeddings).astype("float32")

    # Normalize embeddings for cosine similarity
    faiss.normalize_L2(embeddings)

    dimension = embeddings.shape[1]
    index = faiss.IndexFlatIP(dimension)

    index.add(embeddings)

    return index, embeddings


# -------------------------------
# SEARCH FUNCTION
# -------------------------------
def search_query(query, index, chunks, top_k=3):
    query_embedding = model.encode([query], convert_to_numpy=True)
    query_embedding = np.array(query_embedding).astype("float32")

    faiss.normalize_L2(query_embedding)

    scores, indices = index.search(query_embedding, top_k)

    results = []

    for score, idx in zip(scores[0], indices[0]):
        if idx != -1:
            results.append({
                "chunk": chunks[idx],
                "score": float(score)
            })

    return results


# -------------------------------
# SESSION STATE
# -------------------------------
if "chunks" not in st.session_state:
    st.session_state.chunks = []

if "index" not in st.session_state:
    st.session_state.index = None

if "file_names" not in st.session_state:
    st.session_state.file_names = []


# -------------------------------
# UI HEADER
# -------------------------------
st.title("🌐 Multilingual RAG Assistant")
st.markdown(
    """
    Upload documents in different languages and ask questions in your preferred language.
    
    This first version focuses on **multilingual retrieval**.
    It retrieves the most relevant document chunks even if the document and question are in different languages.
    """
)


# -------------------------------
# SIDEBAR
# -------------------------------
st.sidebar.title("Project Controls")

preferred_language = st.sidebar.selectbox(
    "Preferred Answer Language",
    ["English", "Tamil", "Hindi", "French", "Spanish"]
)

top_k = st.sidebar.slider(
    "Number of chunks to retrieve",
    min_value=1,
    max_value=5,
    value=3
)

st.sidebar.markdown("---")
st.sidebar.info(
    "Current version: Retrieval only. "
    "SLM-based answer generation will be added in the next phase."
)


# -------------------------------
# FILE UPLOAD
# -------------------------------
st.subheader("1. Upload Multilingual Documents")

uploaded_files = st.file_uploader(
    "Upload PDF, TXT, or DOCX files",
    type=["pdf", "txt", "docx"],
    accept_multiple_files=True
)

if uploaded_files:
    if st.button("Process Documents"):
        all_chunks = []
        file_names = []

        with st.spinner("Extracting text and creating multilingual embeddings..."):
            for uploaded_file in uploaded_files:
                text = extract_text(uploaded_file)

                if text.strip():
                    chunks = chunk_text(text)

                    all_chunks.extend(chunks)
                    file_names.append(uploaded_file.name)

            if all_chunks:
                index, embeddings = create_faiss_index(all_chunks)

                st.session_state.chunks = all_chunks
                st.session_state.index = index
                st.session_state.file_names = file_names

                st.success("Documents processed successfully!")

            else:
                st.error("No readable text found in the uploaded documents.")


# -------------------------------
# DOCUMENT STATUS
# -------------------------------
st.subheader("2. Document Status")

if st.session_state.chunks:
    col1, col2, col3 = st.columns(3)

    with col1:
        st.metric("Files Processed", len(st.session_state.file_names))

    with col2:
        st.metric("Text Chunks Created", len(st.session_state.chunks))

    with col3:
        st.metric("Embedding Model", "Multilingual MiniLM")

    with st.expander("View uploaded files"):
        for name in st.session_state.file_names:
            st.write(name)

else:
    st.warning("No documents processed yet.")


# -------------------------------
# QUERY SECTION
# -------------------------------
st.subheader("3. Ask a Question")

query = st.text_input(
    "Enter your question in any language",
    placeholder="Example: What is machine learning? / மெஷின் லெர்னிங் என்றால் என்ன?"
)

if st.button("Search Answer"):
    if not st.session_state.chunks or st.session_state.index is None:
        st.error("Please upload and process documents first.")

    elif not query.strip():
        st.error("Please enter a question.")

    else:
        with st.spinner("Retrieving relevant content..."):
            results = search_query(
                query=query,
                index=st.session_state.index,
                chunks=st.session_state.chunks,
                top_k=top_k
            )

        st.success(f"Retrieved relevant content. Preferred answer language: {preferred_language}")

        st.subheader("Retrieved Context")

        for i, result in enumerate(results, start=1):
            with st.expander(f"Result {i} | Similarity Score: {result['score']:.4f}"):
                st.write(result["chunk"])

        st.subheader("Simple Answer Preview")
        st.info(
            "For now, this preview shows the most relevant retrieved chunk. "
            "In the next phase, we will connect an SLM to generate a proper answer "
            f"in {preferred_language}."
        )

        if results:
            st.write(results[0]["chunk"])


# -------------------------------
# FOOTER
# -------------------------------
st.markdown("---")
st.markdown(
    """
    ### Current MVP Features
    - Upload multilingual documents
    - Extract text from PDF, TXT, and DOCX
    - Split text into chunks
    - Generate multilingual embeddings
    - Store embeddings using FAISS
    - Retrieve relevant chunks for multilingual queries
    
    ### Next Phase
    Add a lightweight SLM to generate final answers in the user's selected language.
    """
)

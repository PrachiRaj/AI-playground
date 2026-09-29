import io
import os

import anthropic
import streamlit as st
from pypdf import PdfReader
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import linear_kernel

ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY", "sk-ant-REPLACE_ME")
MODEL = "claude-opus-5"
CHUNK_WORDS, OVERLAP_WORDS, TOP_K, MAX_HISTORY = 200, 40, 5, 10
NOT_FOUND = "I couldn't find that in the uploaded sources."

SYSTEM = f"""You answer using only the numbered source excerpts in the latest user message.
- Never use outside knowledge.
- Cite every claim with its excerpt number in square brackets, e.g. [2] or [1][3].
- If the excerpts don't contain the answer, reply exactly: "{NOT_FOUND}"
- Citation numbers refer only to the excerpts in the latest message."""

st.set_page_config(page_title="Notebook", page_icon="📚", layout="wide")
client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)


def extract_text(name: str, data: bytes) -> str:
    if name.lower().endswith(".pdf"):
        return "\n".join(p.extract_text() or "" for p in PdfReader(io.BytesIO(data)).pages)
    return data.decode("utf-8", errors="ignore")


def chunk(text: str) -> list[str]:
    words = text.split()
    step = CHUNK_WORDS - OVERLAP_WORDS
    return [" ".join(words[i : i + CHUNK_WORDS]) for i in range(0, max(len(words) - OVERLAP_WORDS, 1), step)] if words else []


@st.cache_resource(show_spinner="Indexing sources…", max_entries=20)
def build_index(files: tuple[tuple[str, bytes], ...]):
    docs = {name: extract_text(name, data) for name, data in files}
    chunks = [(name, c) for name, text in docs.items() for c in chunk(text)]
    try:
        vectorizer = TfidfVectorizer(stop_words="english", ngram_range=(1, 2), sublinear_tf=True)
        matrix = vectorizer.fit_transform(c for _, c in chunks)
    except ValueError:
        return docs, [], None, None
    return docs, chunks, vectorizer, matrix


def retrieve(query: str, chunks, vectorizer, matrix) -> list[tuple[str, str]]:
    scores = linear_kernel(vectorizer.transform([query]), matrix).ravel()
    return [chunks[i] for i in scores.argsort()[::-1][:TOP_K] if scores[i] > 0]


def build_prompt(question: str, items: list[tuple[str, str]]) -> str:
    context = "\n\n".join(f"[{i}] (source: {name})\n{text}" for i, (name, text) in enumerate(items, 1))
    return f"<excerpts>\n{context}\n</excerpts>\n\n{question}"


def stream_answer(messages: list[dict]):
    with client.beta.messages.stream(
        model=MODEL,
        max_tokens=16000,
        system=SYSTEM,
        messages=messages,
        betas=["server-side-fallback-2026-07-01"],
        fallbacks="default",
    ) as stream:
        yield from stream.text_stream
        if stream.get_final_message().stop_reason == "refusal":
            yield "\n\n_The model declined this request._"


def render_sources(items):
    if items:
        with st.expander("Sources"):
            for i, (name, text) in enumerate(items, 1):
                st.markdown(f"**[{i}] {name}**")
                st.caption(text[:300] + ("…" if len(text) > 300 else ""))


def respond(display: str, question: str, items: list[tuple[str, str]]):
    msgs = st.session_state.messages
    msgs.append({"role": "user", "content": display})
    with st.chat_message("user"):
        st.markdown(display)
    with st.chat_message("assistant"):
        if not items:
            answer = NOT_FOUND
            st.markdown(answer)
        else:
            history = [{"role": m["role"], "content": m["content"]} for m in msgs[:-1]][-MAX_HISTORY:]
            try:
                answer = st.write_stream(stream_answer(history + [{"role": "user", "content": build_prompt(question, items)}]))
            except anthropic.APIError as e:
                msgs.pop()
                st.error(f"Claude request failed: {e}")
                return
            render_sources(items)
    msgs.append({"role": "assistant", "content": answer, "sources": items})


st.session_state.setdefault("messages", [])

with st.sidebar:
    st.header("Sources")
    uploads = st.file_uploader("Upload PDF, TXT or MD", type=["pdf", "txt", "md"], accept_multiple_files=True)
    docs, chunks, vectorizer, matrix = build_index(tuple((f.name, f.getvalue()) for f in uploads or []))
    if docs:
        st.caption(f"{len(docs)} file(s) · {len(chunks)} chunks")
    summarize = st.button("Summarize sources", disabled=not chunks)
    if st.button("Clear chat"):
        st.session_state.messages = []
        st.rerun()

st.title("📚 Notebook")
if ANTHROPIC_API_KEY.endswith("REPLACE_ME"):
    st.warning("Set ANTHROPIC_API_KEY at the top of app.py or in the environment.")

for m in st.session_state.messages:
    with st.chat_message(m["role"]):
        st.markdown(m["content"])
        render_sources(m.get("sources"))

if summarize:
    respond(
        "Summarize sources",
        "Summarize each source's key points, then give a short overall synthesis.",
        [(name, text) for name, text in docs.items() if text.strip()],
    )

if prompt := st.chat_input("Ask about your sources", disabled=not chunks):
    respond(prompt, prompt, retrieve(prompt, chunks, vectorizer, matrix))

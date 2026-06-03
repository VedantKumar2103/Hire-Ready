"""
chatbot.py — Gemini-powered AI Career Assistant for HireReady / ParseIQ.

Calls the Gemini REST API directly using `requests` to avoid the httpx
"client has been closed" errors that occur with the google-genai SDK
in Streamlit's rerun execution model.
"""

import streamlit as st
import requests
import json

# ── Gemini REST API endpoints ─────────────────────────────────────────────────

_GEMINI_STREAM_URL = (
    "https://generativelanguage.googleapis.com/v1beta/models/"
    "gemini-2.5-flash:streamGenerateContent"
)

# ── System prompts ────────────────────────────────────────────────────────────

_SYSTEM_PROMPT_WITH_RESUME = """You are **HireReady AI**, an expert career counselor, resume advisor, and interview coach.

The user has uploaded their resume. Here is the extracted information:

- **Name:** {name}
- **Email:** {email}
- **Phone:** {phone}
- **Detected Skills:** {skills}
- **Experience Level:** {cand_level}
- **Resume Score:** {resume_score}/100
- **Recommended Career Field:** {reco_field}
- **Resume Pages:** {no_of_pages}

### Full Resume Text
```
{resume_text}
```

### Your responsibilities:
1. Answer any questions the user has about their resume.
2. Suggest specific, actionable improvements to boost their resume score.
3. Provide tailored career advice for their recommended field ({reco_field}).
4. Help with interview preparation — common questions, STAR method answers, etc.
5. Recommend skills to learn based on gaps you identify.
6. If asked, help rewrite or improve specific resume sections.

Be encouraging, specific, and actionable. Use emojis sparingly for friendliness.
Keep answers concise unless the user asks for detail.
"""

_SYSTEM_PROMPT_NO_RESUME = """You are **HireReady AI**, an expert career counselor, resume advisor, and interview coach.

No resume has been uploaded yet.

### Your responsibilities:
1. Help the user with general career advice and guidance.
2. Share resume writing best practices and tips.
3. Assist with interview preparation — common questions, STAR method, etc.
4. Recommend in-demand skills across various tech fields.
5. Encourage the user to upload their resume in the "📊 Analyze My Resume" tab for personalized, data-driven advice.

Be encouraging, specific, and actionable. Use emojis sparingly for friendliness.
Keep answers concise unless the user asks for detail.
"""


# ── Helpers ───────────────────────────────────────────────────────────────────

def _build_system_prompt(resume_data, resume_text):
    """Return the appropriate system instruction based on available context."""
    if resume_data and resume_text:
        return _SYSTEM_PROMPT_WITH_RESUME.format(
            name=resume_data.get('name', 'N/A'),
            email=resume_data.get('email', 'N/A'),
            phone=resume_data.get('mobile_number', 'N/A'),
            skills=', '.join(resume_data.get('skills', [])),
            cand_level=st.session_state.get('cand_level', 'Unknown'),
            resume_score=st.session_state.get('resume_score', '?'),
            reco_field=st.session_state.get('reco_field', 'Not determined yet'),
            no_of_pages=resume_data.get('no_of_pages', '?'),
            resume_text=resume_text[:8000],
        )
    return _SYSTEM_PROMPT_NO_RESUME


def _build_api_contents(chat_messages):
    """
    Convert our chat_messages list into the Gemini REST API content format.
    """
    contents = []
    for msg in chat_messages:
        role = "user" if msg["role"] == "user" else "model"
        contents.append({
            "role": role,
            "parts": [{"text": msg["content"]}],
        })
    return contents


def _stream_gemini(api_key, system_prompt, chat_messages, new_prompt):
    """
    Call the Gemini REST API with streaming using `requests`.
    Yields text chunks as they arrive via Server-Sent Events.
    """
    url = f"{_GEMINI_STREAM_URL}?key={api_key}&alt=sse"

    # Build request body
    contents = _build_api_contents(chat_messages)
    contents.append({
        "role": "user",
        "parts": [{"text": new_prompt}],
    })

    body = {
        "system_instruction": {"parts": [{"text": system_prompt}]},
        "contents": contents,
    }

    # Make streaming request (requests library — no httpx lifecycle issues)
    resp = requests.post(url, json=body, stream=True, timeout=120)

    if resp.status_code != 200:
        error_detail = resp.text[:500]
        raise RuntimeError(f"Gemini API error ({resp.status_code}): {error_detail}")

    # Parse Server-Sent Events
    for line in resp.iter_lines(decode_unicode=True):
        if line and line.startswith("data: "):
            data_str = line[6:]
            try:
                data = json.loads(data_str)
                candidates = data.get("candidates", [])
                if candidates:
                    parts = candidates[0].get("content", {}).get("parts", [])
                    for part in parts:
                        text = part.get("text", "")
                        if text:
                            yield text
            except json.JSONDecodeError:
                continue


# ── Main UI renderer ─────────────────────────────────────────────────────────

def render_chatbot(resume_data=None, resume_text=None):
    """
    Render the AI Career Assistant chat interface.

    Parameters
    ----------
    resume_data : dict or None
        Parsed resume dict (name, email, skills, etc.) from session_state.
    resume_text : str or None
        Raw text extracted from the uploaded resume PDF.
    """

    st.header("🤖 AI Career Assistant")
    st.caption("Powered by Google Gemini — ask me anything about your career, resume, or interviews!")

    # ── API key (pre-set in session_state by App9.py) ────────────────
    api_key = st.session_state.get('gemini_api_key', '')
    if not api_key:
        st.error("⚠️ Gemini API key not found. Please restart the app.")
        return

    # ── Context badge ─────────────────────────────────────────────────
    if resume_data:
        st.success(
            f"✅ Resume loaded for **{resume_data.get('name', 'User')}** — "
            f"I can give you personalized advice!"
        )
    else:
        st.warning(
            "📄 No resume uploaded yet. Upload one in the **📊 Analyze My Resume** tab "
            "for personalized advice, or ask me anything about careers!"
        )

    # ── System prompt ─────────────────────────────────────────────────
    system_prompt = _build_system_prompt(resume_data, resume_text)

    # ── Chat history (stored as plain dicts — always serialisable) ────
    if 'chat_messages' not in st.session_state:
        st.session_state.chat_messages = []

    # Render previous messages
    for msg in st.session_state.chat_messages:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    # ── User input ────────────────────────────────────────────────────
    if prompt := st.chat_input("Ask me anything about your career…"):
        # Show user message
        st.session_state.chat_messages.append({"role": "user", "content": prompt})
        with st.chat_message("user"):
            st.markdown(prompt)

        # ── Stream from Gemini REST API ───────────────────────────────
        with st.chat_message("assistant"):
            placeholder = st.empty()
            full_response = ""
            try:
                for text_chunk in _stream_gemini(
                    api_key,
                    system_prompt,
                    st.session_state.chat_messages[:-1],  # history before this prompt
                    prompt,
                ):
                    full_response += text_chunk
                    placeholder.markdown(full_response + " ▌")

                placeholder.markdown(full_response)

            except Exception as e:
                error_msg = f"⚠️ Sorry, something went wrong: {e}"
                placeholder.markdown(error_msg)
                full_response = error_msg

        # Save assistant reply
        st.session_state.chat_messages.append(
            {"role": "assistant", "content": full_response}
        )

    # ── Quick-action buttons (first visit only) ──────────────────────
    if not st.session_state.chat_messages:
        st.markdown("---")
        st.markdown("#### 💡 Try asking:")
        cols = st.columns(2)
        if resume_data:
            suggestions = [
                "How can I improve my resume score?",
                "What skills should I learn next?",
                "Help me prepare for an interview",
                "Rewrite my resume summary",
            ]
        else:
            suggestions = [
                "What makes a great resume?",
                "Top skills for software engineers in 2026",
                "Common interview mistakes to avoid",
                "How to write a strong cover letter",
            ]
        for i, suggestion in enumerate(suggestions):
            with cols[i % 2]:
                if st.button(f"💬 {suggestion}", key=f"suggest_{i}", use_container_width=True):
                    st.session_state.chat_messages.append(
                        {"role": "user", "content": suggestion}
                    )
                    st.rerun()
<div align="center">

# 🎯 Hire Ready

### Your AI-Powered Career Companion

*Analyze. Build. Get Hired.*

[![Python](https://img.shields.io/badge/Python-3.10+-blue.svg)](https://www.python.org/)
[![Gemini](https://img.shields.io/badge/Powered_by-Gemini_AI-4285F4.svg)](https://ai.google.dev/)
[![n8n](https://img.shields.io/badge/Automated_with-n8n-EA4B71.svg)](https://n8n.io/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

[Demo](#-demo) • [Features](#-what-can-it-do) • [Quick Start](#-quick-start) • [Architecture](#-architecture) • [Roadmap](#-roadmap)

</div>

---

## 🌟 The Problem

Every year, **75% of resumes never reach a human recruiter** — they're filtered out by Applicant Tracking Systems (ATS). Job seekers waste hours rewriting resumes, guessing what works, and struggling to match job descriptions.

**Hire Ready solves this.**

---

## 💡 The Solution

A one-stop platform that:
1. 🔍 **Scans** your existing resume and tells you exactly what's wrong
2. ✏️ **Builds** a brand-new ATS-friendly resume from scratch
3. 💬 **Coaches** you with an AI career mentor (Gemini-powered)
4. 📬 **Delivers** detailed feedback straight to your inbox

---

## 🎬 Demo

> 📹 _Add your demo video or GIF here_

```
🎥 demo.mp4   |   🖼️ screenshots/
```

---

## ⚡ What Can It Do?

<table>
<tr>
<td width="50%" valign="top">

### 🔍 Resume Analyzer
Drop in your PDF or DOCX resume. Within seconds, get:
- 📊 ATS-readiness score
- 🎯 Missing keywords for your target role
- ⚠️ Formatting red flags
- ✨ Section-by-section critique

</td>
<td width="50%" valign="top">

### 🏗️ Smart Resume Builder
No design skills needed. Just fill the form:
- 👤 Personal details & contact
- 🎓 Education & certifications
- 💼 Work experience
- 🛠️ Skills & projects
→ Outputs a **clean, ATS-friendly PDF**

</td>
</tr>
<tr>
<td width="50%" valign="top">

### 🤖 AI Career Mentor
Chat with a Gemini-powered assistant trained for career guidance:
- "How do I prepare for a Data Analyst interview?"
- "What skills should I add for Frontend roles?"
- "Review this project description"

</td>
<td width="50%" valign="top">

### 📧 Auto-Delivered Reports
Hit one button — your full report lands in your inbox:
- ✉️ Powered by **n8n workflow**
- 📄 Professional PDF format
- 🔁 Share with mentors, friends, recruiters

</td>
</tr>
</table>

---

## 🛠️ Built With

```
🐍 Backend       →  Python • Streamlit / Flask
🧠 AI Engine     →  Google Gemini API
📄 Parsing       →  PyPDF2 • pdfminer • python-docx
⚙️ Automation    →  n8n (webhook + email node)
🎨 Frontend      →  HTML • CSS • JavaScript
📦 PDF Engine    →  ReportLab • fpdf2
🚀 Deployment    →  Render / Streamlit Cloud / Vercel
```

---

## 🏗️ Architecture

```
                    ┌──────────────────┐
                    │   USER UPLOADS   │
                    │  (PDF / DOCX)    │
                    └────────┬─────────┘
                             │
                             ▼
         ┌───────────────────────────────────┐
         │       RESUME PARSER MODULE        │
         │   (extracts text, structure)      │
         └───────────────┬───────────────────┘
                         │
                         ▼
         ┌───────────────────────────────────┐
         │      GEMINI AI ANALYZER           │
         │  • Scores  • Gaps  • Suggestions  │
         └───────┬───────────────────┬───────┘
                 │                   │
                 ▼                   ▼
    ┌────────────────────┐  ┌──────────────────────┐
    │  ON-SCREEN REPORT  │  │   n8n WEBHOOK CALL   │
    │   (user sees it)   │  │ → SMTP → User Inbox  │
    └────────────────────┘  └──────────────────────┘
```

---

## 🚀 Quick Start

### 📋 Prerequisites
| Requirement | Where to get it |
|-------------|-----------------|
| Python 3.10+ | [python.org](https://www.python.org/downloads/) |
| Gemini API Key | [Google AI Studio](https://aistudio.google.com/app/apikey) |
| n8n Instance | [n8n.io](https://n8n.io/) (cloud or self-host) |

### 🔧 Installation

```bash
# 1. Clone the repository
git clone https://github.com/VedantKumar2103/hire-ready.git
cd hire-ready

# 2. Create & activate virtual environment
python -m venv venv
source venv/bin/activate          # Linux / Mac
venv\Scripts\activate             # Windows

# 3. Install dependencies
pip install -r requirements.txt

# 4. Configure your secrets
cp .env.example .env
# Open .env and paste your Gemini API key + n8n webhook URL
```

### ▶️ Run the App

```bash
# If using Streamlit
streamlit run app.py

# If using Flask
python app.py
```

🌐 Open `http://localhost:8501` (Streamlit) or `http://localhost:5000` (Flask)

---

## 🔐 Environment Variables

Create a `.env` file in the root (use `.env.example` as a template):

```env
GEMINI_API_KEY=your_gemini_key_here
N8N_WEBHOOK_URL=https://your-n8n.com/webhook/xxx
```

> ⚠️ **Never commit your `.env` file.** It's in `.gitignore` for a reason.

---

## 🔄 n8n Workflow Setup

The email automation runs through n8n. Setup in 3 steps:

1. **Import the workflow** — load `n8n/workflow.json` into your n8n canvas
2. **Connect email credentials** — Gmail OAuth / SMTP / SendGrid (your choice)
3. **Copy the webhook URL** into your `.env` as `N8N_WEBHOOK_URL`

That's it. Every analysis request now auto-emails the user.

---

## 📂 Project Layout

```
hire-ready/
│
├── 📄 app.py                    # Application entry point
├── 📋 requirements.txt          # Python dependencies
├── 🔐 .env.example              # Env template
├── 🚫 .gitignore                # Git ignore rules
├── 📖 README.md                 # You are here
│
├── 📁 modules/
│   ├── analyzer.py              # Resume parsing + Gemini analysis
│   ├── builder.py               # ATS resume builder logic
│   ├── chatbot.py               # Gemini chatbot wrapper
│   └── notifier.py              # n8n webhook trigger
│
├── 📁 templates/                # Jinja2 / HTML templates
├── 📁 static/                   # CSS, JS, assets
├── 📁 n8n/
│   └── workflow.json            # n8n workflow export
│
└── 📁 screenshots/              # App preview images
```

---

## 🎯 Who Is This For?

| Persona | Why they'll love it |
|---------|---------------------|
| 🎓 **Students** | Build a first-ever resume that actually passes ATS |
| 🔄 **Career switchers** | Tailor your resume to a new domain in minutes |
| 💼 **Active job seekers** | Stop guessing — get data-driven feedback |
| 👨‍🏫 **Career coaches** | Automate the first round of resume reviews |
| 🏢 **Recruiters** | Pre-screen candidates faster |

---

## 🗺️ Roadmap

- [x] PDF / DOCX resume parsing
- [x] Gemini-powered analysis engine
- [x] ATS-friendly resume builder
- [x] AI career chatbot
- [x] n8n email automation
- [ ] LinkedIn profile import
- [ ] Job description ↔ resume matcher with score
- [ ] Cover letter generator
- [ ] Mock interview module
- [ ] Multi-language support (Hindi, Spanish, French)
- [ ] Browser extension version

---

## 🐛 Found a Bug?

Open an [issue](https://github.com/VedantKumar2103/hire-ready/issues) — please include:
- What you did
- What you expected
- What actually happened
- Screenshots (if applicable)

---

## 🤝 Want to Contribute?

PRs welcome! Here's the flow:

```bash
1. Fork this repo
2. git checkout -b feature/amazing-thing
3. git commit -m "Add amazing thing"
4. git push origin feature/amazing-thing
5. Open a Pull Request
```

---

## 📜 License

Released under the **MIT License** — free to use, modify, and distribute.

---

# Hire Ready — AI Resume Analyzer & Builder

A full-stack web application that helps job seekers create, analyze, and improve their resumes using AI. Hire Ready combines a **resume analyzer**, an **ATS-friendly resume builder**, and a **Gemini-powered career chatbot** — all in one platform, with automated email delivery powered by **n8n workflow automation**.

---

## ✨ Features

### 📄 1. Resume Analyzer
- Upload your existing resume (PDF/DOCX)
- Receive a detailed **resume report** with:
  - ATS compatibility score
  - Skill-gap analysis
  - Section-wise feedback
  - Improvement suggestions
- Get the **full report delivered to your email** via n8n workflow automation

### 📝 2. ATS-Friendly Resume Builder
- Fill in a structured form (personal details, education, skills, experience, projects)
- Generate a clean, **ATS-compliant resume** instantly
- Download as PDF — ready to send to recruiters

### 🤖 3. AI Career Chatbot
- Powered by **Google Gemini API**
- Get personalized career guidance based on your target role
- Ask about interview prep, skill roadmaps, and industry trends
- Conversational interface for natural Q&A

### 📧 4. Email Report Delivery
- Send your resume analysis report to your inbox
- Automated through **n8n workflow** — no manual export needed
- Shareable, professional PDF format

---

## 🛠 Tech Stack

| Layer | Technologies |
|-------|-------------|
| **Frontend** | HTML, CSS, JavaScript |
| **Backend** | Python (Flask / Streamlit) |
| **AI** | Google Gemini API |
| **Resume Parsing** | PyPDF2 / pdfminer / python-docx |
| **Automation** | n8n workflow (email delivery) |
| **Deployment** | (e.g., Render / Vercel / Streamlit Cloud) |

---

## 📁 Project Structure

```
hire-ready/
├── app.py                      # Main application entry point
├── requirements.txt
├── .env.example                # Template for environment variables
├── README.md
│
├── modules/
│   ├── analyzer.py             # Resume analysis logic
│   ├── builder.py              # ATS resume builder
│   ├── chatbot.py              # Gemini API integration
│   └── email_service.py        # n8n webhook trigger
│
├── templates/                  # HTML templates
├── static/                     # CSS, JS, images
└── n8n/
    └── workflow.json           # n8n workflow export
```

---

## 🚀 Getting Started

### Prerequisites
- Python 3.10+
- Google Gemini API key — [get one here](https://aistudio.google.com/app/apikey)
- n8n instance (cloud or self-hosted) for email automation

### Installation

```bash
# Clone the repo
git clone https://github.com/YOUR_USERNAME/hire-ready.git
cd hire-ready

# Create virtual environment
python -m venv venv
source venv/bin/activate   # On Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Configure environment variables
cp .env.example .env
# Edit .env with your Gemini API key and n8n webhook URL
```

### Environment Variables

Create a `.env` file with:

```env
GEMINI_API_KEY=your_gemini_api_key_here
N8N_WEBHOOK_URL=https://your-n8n-instance.com/webhook/xxx
```

### Run the App

```bash
python app.py
# or for Streamlit:
streamlit run app.py
```

Open your browser at `http://localhost:5000` (Flask) or `http://localhost:8501` (Streamlit).

---

## 🔄 n8n Workflow Setup

1. Import the workflow file from `n8n/workflow.json` into your n8n instance.
2. Configure the email node (SMTP / Gmail / SendGrid credentials).
3. Copy the webhook URL into your `.env` as `N8N_WEBHOOK_URL`.
4. Activate the workflow.

The workflow listens for resume analysis events and emails the report to the user automatically.

---

## 💡 How It Works

```
┌────────────────┐     ┌──────────────────┐     ┌─────────────────┐
│  User uploads  │ ──▶ │  Resume Parser   │ ──▶ │  Gemini-based   │
│     resume     │     │  (PDF / DOCX)    │     │   Analyzer      │
└────────────────┘     └──────────────────┘     └────────┬────────┘
                                                          │
┌────────────────┐     ┌──────────────────┐               │
│   User email   │ ◀── │   n8n Webhook    │ ◀────────────┘
│    inbox       │     │   (Email node)   │     PDF report
└────────────────┘     └──────────────────┘
```

---

## 📸 Screenshots

> _(Add screenshots of your dashboard, builder, chatbot, and email here)_

- `screenshots/dashboard.png`
- `screenshots/resume-builder.png`
- `screenshots/chatbot.png`
- `screenshots/email-report.png`

---

## 🎯 Use Cases

- **Students** preparing for placements
- **Job seekers** transitioning between roles
- **Recruiters** screening candidates faster
- **Career coaches** offering automated feedback

---

## 🔮 Future Enhancements

- [ ] LinkedIn profile import
- [ ] Cover letter generator
- [ ] Multi-language support
- [ ] Mock interview simulator
- [ ] Job description matcher with score

---

## 🤝 Contributing

Contributions are welcome! Feel free to open issues or submit pull requests.

```bash
# Fork the repo
# Create a branch
git checkout -b feature/your-feature-name

# Commit your changes
git commit -m "Add some feature"

# Push and open a PR
git push origin feature/your-feature-name
```

---

## 📜 License

MIT License — feel free to use and adapt.

---

## 👤 Author

**Vedant Kumar**
- 📧 vedantkumar0411@gmail.com
- 🔗 [LinkedIn](https://linkedin.com/in/vedant-kumar-909b2626a)
- 💻 [GitHub](https://github.com/VedantKumar2103)

---

⭐ **If you found this project helpful, please give it a star on GitHub!**

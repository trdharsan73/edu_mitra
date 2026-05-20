# EduMitra: Multilingual AI Tutor

EduMitra is a web-based AI tutor designed to help school and college students understand concepts, solve doubts, and practice in English and Tamil (with code-mixing support). It focuses on CBSE/State Board subjects like Mathematics, Physics, Chemistry, and Biology (Classes 8–12).

## Key Features
- **Step-by-step explanations**
- **Personalized responses** based on user level
- **Practice questions with evaluation**
- **Grounded answers** using NCERT/State Board content (to reduce hallucinations) via Retrieval-Augmented Generation (RAG)
- **Multilingual** (English + Tamil support)

## Tech Stack
- **Frontend**: Streamlit
- **LLM**: Google Gemini API (or Groq Llama 3.1)
- **RAG Framework**: LangChain
- **Vector DB**: ChromaDB
- **Embeddings**: HuggingFace Sentence Transformers (`all-MiniLM-L6-v2`)
- **PDF Extraction**: PyMuPDF

## Phased Development Plan
- [x] Phase 0: Setup & Planning (Downloaded PDF data)
- [ ] Phase 1: Core RAG Pipeline
- [ ] Phase 2: AI Tutor Logic & Prompt Engineering
- [ ] Phase 3: Web Interface & Features (Streamlit)
- [ ] Phase 4: Multilingual Improvements & Polish
- [ ] Phase 5: WhatsApp Integration (Optional)
- [ ] Phase 6: Testing, Documentation & Deployment
- [ ] Phase 7: Bonus & Presentation

## Setup Instructions

1. **Create Virtual Environment**
   ```bash
   python -m venv venv
   .\venv\Scripts\activate
   ```

2. **Install Dependencies**
   ```bash
   pip install -r requirements.txt
   ```

3. **Environment Variables**
   Create a `.env` file in the root directory and add your API keys:
   ```env
   GEMINI_API_KEY=your_gemini_api_key
   GROQ_API_KEY=your_groq_api_key
   ```

4. **Run Application**
   ```bash
   streamlit run src/app.py
   ```

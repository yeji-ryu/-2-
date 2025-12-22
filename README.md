# 🛡️ AI-Based Confidential Investment Report Similarity & Leakage Detection System  

A research & engineering project focused on preventing **internal investment document leakage** using **AI-driven document similarity analysis**.  
The system detects confidential content reuse by comparing newly uploaded reports with internally stored analyst documents using hybrid scoring (LLM + BM25 + Embedding cosine similarity).  

---

## 🌐 Notion Project Space  
🔗 https://www.notion.so/2-26bcf407bc8880e38e86e00acd031929?source=copy_link  
This space documents early project planning, brainstorming notes, architecture drafts, and concept design.

---

## 📌 Project Objective  

In the financial industry, a single leaked forecast number, valuation metric, or confidential chart can manipulate market behavior and cause enormous socioeconomic damage.  
Cases such as the **SG Securities market crash** and **CJ ENM stock manipulation scandal** prove the real-world impact of internal data misuse.  

This project builds a system to:  

- Embed internal investment reports into vector storage  
- Compare new documents against internal data using AI  
- Detect meaningful similarity beyond keyword matching  
- Identify potential confidential information leakage  
- Provide automated similarity summaries and rankings  

The goal is to **catch internal information reuse before publication**.

---

## 🚀 Key Innovation vs Traditional DLP Systems  

Traditional DLP systems rely heavily on patterns, keywords, and regex → detecting only literal matches.  
Our system:

| Property | Old DLP | Generic AI Similarity | Proposed System |
|---------|--------|----------------------|----------------|
| Method | Keyword / Regex | Embedding or BM25 | Hybrid Embedding + BM25 + LLM |
| Contextual Understanding | Very Low | Medium | High |
| Domain Adapted | No | Partially | Fully industry-specific |
| Accuracy Boost | Manual rules | Statistical weights | LLM scoring refinement |
| Output | True/False | Similarity scores | Context-aware explanations |

Our model is designed specifically for **investment reporting language, structure, terminology and writing patterns.**

---

## 📂 Repository Structure  

root
├── corpus_docs/ # Internal confidential investment analyst reports (~30 docs, each ~30 pages)
├── external_docs/ # External documents, diverse domains and topics
│
├── v1_folder/ # Early prototype version
├── v2_folder/ # Second iteration with improved pipeline
├── final_version_folder/ # Final optimized model + all final notebooks and scripts
│
├── document/ # Progress reports + festival workbook + weekly documentation
│
├── fake_report_gemini.py # LLM-based automated internal document generation script
├── README.md
└── video/ # Website demo + system visualization recording

yaml
코드 복사

---

## 🧠 System Architecture

The system processes documents through **two independent similarity paths** and merges them using LLM evaluation:

### 🔹 Embedding Path (Cosine Similarity)
1. Chunk documents into meaningful blocks  
2. Embed using `bge-m3-ko`
3. Store vectors in Qdrant
4. Extract cosine similarity top-n results

→ Measures semantic similarity beyond vocabulary overlap  

---

### 🔹 BM25 Sparse Vector Path
1. Tokenize & normalize  
2. Push into BM25 ranking index  
3. Extract TF-IDF based similarity  

→ Measures keyword & term frequency structural similarity  

---

### 🔹 LLM Hybrid Decision Layer  
LLM reads both scoring structures extracted from:

- cosine similarity  
- BM25 ranking  
- percentile distributions  
- trim scoring patterns  

and **produces a dynamic hybrid score + explanation**.

---

## 📄 Output Example (Single Markdown Page)

Document Similarity Analysis Result
Input Target
Uploaded document: 누누푼토리 투자보고서 (external)

Top Similar Internal Documents (Hybrid Score Ranking)
1️⃣ 시그마이더투자보고서 — Hybrid: 48.38
2️⃣ 라온투자보고서 — Hybrid: 43.22
3️⃣ 스카이라인투자보고서 — Hybrid: 35.46

Cosine Similarity Highlights
Sigma battery report: 0.22

Future battery report: 0.24

Medical Pharma 2024: 0.19

BM25 Highlights
Sigma battery: 57.29

Future battery: 52.12

Skyline medical: 51.81

Extracted Similar Sentences (Semantic Match)
“기술특허 관련 SoC 자산 세부 영향과 게이트웨이 리스크 요건 분석 내용이 포함된 보고서”

“핵심 동향 분석: 5년 IRR 예측 기반 투자수익 분석 수치 포함”

Summary Interpretation
Similar industry category (Battery & industrial tech)

Overlapping sentence structure and concept alignment

High BM25 keyword locality match

IRR / CAGR / forecast metrics similar

yaml
코드 복사

→ This page is automatically generated for each upload.

---

## 🎥 Web Demo Video  

A running webpage prototype demonstrating:

- Document upload UI  
- Automated report generation  
- Similarity computation visualization  
- Ranked document results  

📁 video/ folder contains:  
- live system demo  
- web interface workflow  

---

## 🔬 Models + Algorithms Used

| Component | Implementation |
|----------|---------------|
| Embedding Model | `bge-m3-ko` |
| Sparse Retrieval | BM25 (rank_bm25) |
| Vector DB | Qdrant |
| Hybrid Scoring | LLM-based evaluation |
| Similarity Output | Markdown auto-generation |

---

## 🏗️ Development Stack

- Python, FastAPI  
- Qdrant, BM25  
- LangChain, ChatGPT/Gemini  
- Docker / GitHub Actions  
- HTML/JS web demo  
- GitHub + Notion collaboration  

---

## 📌 Future Improvements  

- Multilingual support  
- OCR table/graph extraction  
- Layout-aware embeddings  
- multimodal similarity analysis  
- financial ontology knowledge graph  

---

## 👤 Team Members

| Name     | Major | Role |
|----------|------|-----|
| Yeji Ryu | CS + Security | NLP / Modeling |
| Donggyun Han | CS + Security | Backend / Architecture |

---

## 📄 Academic Context

University course project:  
**Integrated Security Pre-Capstone Design — Industry Collaboration Track**

Partner company:  
**SOMANSA – Data Privacy & Enterprise Security**

---

Thanks for reading this repository!  
Feel free to explore the codebase and research results.  

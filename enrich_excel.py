"""
Enrich INT Diary Entries xlsx:
- Fill missing module/syllabus/practical for entries 41-84
- Generate proper Work Summary (2-3 sentences) & Learnings (4-5 numbered points)
  for ALL entries, similar to the first day's style
- Write enriched content back into the Excel columns G (Work Summary) and H (Learnings)
"""

import openpyxl
from datetime import datetime
from copy import copy

EXCEL_FILE = "INT Diary Entries_2026-27.xlsx"

# ─────────────────────────────────────────────────────────────────
# TOPICS FOR EMPTY ENTRIES (41 onwards, Oct 6 → Oct 31 for portal)
# ─────────────────────────────────────────────────────────────────
# Keyed by 1-based entry index (matching row order in sheet)
EXTRA_TOPICS = {
    41: {
        "module": "Fine-tuning LLMs",
        "syllabus": "Introduction to fine-tuning; task-specific adaptation; full fine-tuning vs parameter-efficient methods",
        "practical": "Identify a use case for fine-tuning a pre-trained LLM",
    },
    42: {
        "module": "Fine-tuning LLMs",
        "syllabus": "LoRA and QLoRA; low-rank adaptation; adapter layers; PEFT library overview",
        "practical": "Apply LoRA to adapt a transformer model on a small dataset",
    },
    43: {
        "module": "Prompt Engineering",
        "syllabus": "Zero-shot prompting; few-shot prompting; chain-of-thought reasoning; prompt templates",
        "practical": "Design and test prompts for a text classification task",
    },
    44: {
        "module": "Prompt Engineering",
        "syllabus": "Advanced prompting strategies; self-consistency; tree-of-thought; role prompting",
        "practical": "Compare prompt strategies on a reasoning benchmark",
    },
    45: {
        "module": "RAG Foundations",
        "syllabus": "Retrieval-Augmented Generation overview; motivation; architecture; retrieval vs parametric memory",
        "practical": "Design a RAG pipeline architecture diagram",
    },
    46: {
        "module": "RAG Implementation",
        "syllabus": "Document chunking; embedding models; similarity search; vector retrieval",
        "practical": "Chunk documents and compute embeddings using sentence-transformers",
    },
    47: {
        "module": "Vector Databases",
        "syllabus": "FAISS; Chroma; Pinecone; indexing strategies; approximate nearest-neighbor search",
        "practical": "Build a simple vector store and query it with a natural language question",
    },
    48: {
        "module": "LangChain Basics",
        "syllabus": "LangChain framework; chains; prompts; LLMs; output parsers; document loaders",
        "practical": "Build a basic Q&A chain using LangChain and an open LLM",
    },
    49: {
        "module": "LangChain Advanced",
        "syllabus": "Memory in LangChain; conversation buffer; retrieval chains; RetrievalQA",
        "practical": "Add conversational memory to a LangChain RAG chain",
    },
    50: {
        "module": "AI Agents",
        "syllabus": "Agent architectures; ReAct framework; tool use; planning and reasoning loops",
        "practical": "Build a simple agent that uses tools to answer multi-step questions",
    },
    51: {
        "module": "Model Deployment",
        "syllabus": "Model serving; FastAPI; REST API design for ML; request/response schemas",
        "practical": "Wrap a trained model in a FastAPI endpoint",
    },
    52: {
        "module": "Containerization",
        "syllabus": "Docker basics; Dockerfile; containerizing ML applications; Docker Compose",
        "practical": "Containerize the FastAPI ML service with Docker",
    },
    53: {
        "module": "MLOps Basics",
        "syllabus": "MLOps principles; experiment tracking; MLflow; model versioning and registry",
        "practical": "Log experiments and metrics using MLflow",
    },
    54: {
        "module": "End-to-End AI Project",
        "syllabus": "Project architecture; integrating RAG + LLM + vector DB into a full pipeline",
        "practical": "Build and test the end-to-end AI application pipeline",
    },
    55: {
        "module": "Frontend Integration",
        "syllabus": "React basics for AI apps; API calls from frontend; state management; UI/UX for AI",
        "practical": "Build a React frontend that communicates with the AI backend",
    },
    56: {
        "module": "Testing AI Applications",
        "syllabus": "Unit testing; integration testing; evaluating LLM outputs; testing RAG pipelines",
        "practical": "Write test cases for the AI application components",
    },
    57: {
        "module": "Performance Optimization",
        "syllabus": "Inference optimization; quantization; batching; caching; response latency reduction",
        "practical": "Profile and optimize the deployed model's inference speed",
    },
    58: {
        "module": "Project Documentation",
        "syllabus": "Technical writing; API documentation; README structure; architecture diagrams",
        "practical": "Write complete API documentation and a detailed project README",
    },
    59: {
        "module": "Code Review & Refactoring",
        "syllabus": "Code quality principles; refactoring techniques; clean code; review best practices",
        "practical": "Conduct a code review and refactor the project codebase",
    },
    60: {
        "module": "Project Presentation & Summary",
        "syllabus": "Presenting AI projects; demo preparation; summarizing learnings; future improvements",
        "practical": "Prepare and deliver a final project demonstration and summary report",
    },
    # Nov entries (61–84) — additional AI/ML topics
    61: {
        "module": "Responsible AI",
        "syllabus": "AI ethics; bias and fairness; explainability; accountability; responsible AI frameworks",
        "practical": "Analyze a model for bias and document mitigation strategies",
    },
    62: {
        "module": "Explainable AI",
        "syllabus": "SHAP values; LIME; feature importance; model interpretability techniques",
        "practical": "Apply SHAP to explain predictions of a trained model",
    },
    63: {
        "module": "Multimodal AI",
        "syllabus": "Multimodal models; vision-language models; CLIP; image-text alignment",
        "practical": "Experiment with a CLIP model for image-text retrieval",
    },
    64: {
        "module": "Speech & Audio AI",
        "syllabus": "Speech recognition; Whisper; text-to-speech; audio feature extraction",
        "practical": "Transcribe audio using OpenAI Whisper",
    },
    65: {
        "module": "Reinforcement Learning from Human Feedback",
        "syllabus": "RLHF overview; reward modeling; PPO for LLM alignment; Constitutional AI",
        "practical": "Study RLHF pipeline and document its role in LLM alignment",
    },
    66: {
        "module": "Knowledge Graphs",
        "syllabus": "Knowledge graphs; triples; ontologies; graph databases; GraphRAG",
        "practical": "Build a small knowledge graph and query it with SPARQL",
    },
    67: {
        "module": "AI Safety",
        "syllabus": "AI safety principles; hallucination mitigation; guardrails; red-teaming LLMs",
        "practical": "Red-team a language model and document failure modes",
    },
    68: {
        "module": "AutoML",
        "syllabus": "AutoML concepts; hyperparameter search; neural architecture search; AutoGluon",
        "practical": "Run an AutoML experiment on a classification dataset",
    },
    69: {
        "module": "Data Engineering for AI",
        "syllabus": "Data pipelines; ETL; data versioning with DVC; feature stores",
        "practical": "Build a data pipeline and version datasets using DVC",
    },
    70: {
        "module": "Cloud AI Services",
        "syllabus": "Google Vertex AI; AWS SageMaker; Azure AI; managed ML services overview",
        "practical": "Deploy a model on a cloud ML platform and test inference",
    },
    71: {
        "module": "LLM Evaluation",
        "syllabus": "LLM benchmarks; BLEU, ROUGE, BERTScore; human evaluation; LLM-as-judge",
        "practical": "Evaluate a fine-tuned model using standard NLP metrics",
    },
    72: {
        "module": "Generative AI Applications",
        "syllabus": "Real-world GenAI applications; chatbots; code generation; content creation; summarization",
        "practical": "Build a document summarization tool using a generative model",
    },
    73: {
        "module": "Federated Learning",
        "syllabus": "Federated learning overview; privacy-preserving ML; differential privacy; secure aggregation",
        "practical": "Simulate a federated learning setup across multiple clients",
    },
    74: {
        "module": "On-device AI",
        "syllabus": "Edge AI; model compression; pruning; quantization for mobile; TFLite and ONNX",
        "practical": "Convert a model to TFLite and run inference on a simulated edge device",
    },
    75: {
        "module": "AI Project Management",
        "syllabus": "Agile for AI; sprint planning; AI project lifecycle; stakeholder communication",
        "practical": "Create a sprint plan for a hypothetical AI product feature",
    },
    76: {
        "module": "AI in Industry",
        "syllabus": "AI use cases in healthcare, finance, education and manufacturing; case studies",
        "practical": "Research and present an AI case study from a chosen industry",
    },
    77: {
        "module": "Capstone Project Planning",
        "syllabus": "Defining problem statement; dataset selection; model selection; evaluation criteria",
        "practical": "Write the project proposal for the capstone AI project",
    },
    78: {
        "module": "Capstone Project Development",
        "syllabus": "Data preprocessing; model training; baseline results; iterative improvement",
        "practical": "Train and evaluate the baseline model for the capstone project",
    },
    79: {
        "module": "Capstone Project Optimization",
        "syllabus": "Hyperparameter tuning; ensembling; error analysis; model improvement strategies",
        "practical": "Optimize the capstone model and document performance gains",
    },
    80: {
        "module": "Capstone Frontend & API",
        "syllabus": "Integrating trained model into a web app; REST API; user interface design",
        "practical": "Build and test the full-stack interface for the capstone project",
    },
    81: {
        "module": "Capstone Testing & Validation",
        "syllabus": "End-to-end testing; user acceptance testing; performance benchmarking",
        "practical": "Conduct end-to-end testing and fix identified issues",
    },
    82: {
        "module": "Capstone Documentation",
        "syllabus": "Final report writing; methodology documentation; results presentation",
        "practical": "Write the final internship project report",
    },
    83: {
        "module": "Internship Review",
        "syllabus": "Reviewing all topics covered; skill assessment; identifying strengths and gaps",
        "practical": "Create a personal learning summary document for the internship period",
    },
    84: {
        "module": "Final Submission & Closure",
        "syllabus": "Final code submission; repository cleanup; handover documentation; internship closure",
        "practical": "Submit all deliverables and complete the internship closure checklist",
    },
}

# ─────────────────────────────────────────────────────────────────
# CONTENT GENERATOR
# ─────────────────────────────────────────────────────────────────

def generate_work_summary(module: str, syllabus: str, practical: str) -> str:
    return (
        f"{module} involves studying and applying key concepts that are central to "
        f"modern AI and Machine Learning development. "
        f"During this session, the focus was on: {syllabus}. "
        f"Hands-on work included the following practical activity: {practical}."
    )


def generate_learnings(module: str, syllabus: str, practical: str) -> str:
    topics = [t.strip() for t in syllabus.split(";") if t.strip()]
    lines = []
    for i, topic in enumerate(topics[:5], start=1):
        lines.append(f"{i}. Understood and applied the concept of {topic} as part of the {module} module.")

    # Pad to at least 4 points
    if len(lines) < 4:
        lines.append(f"{len(lines)+1}. Completed the practical task: {practical}, which reinforced theoretical understanding.")
    if len(lines) < 4:
        lines.append(f"{len(lines)+1}. Developed problem-solving skills related to {module} through hands-on exercises.")
    if len(lines) < 4:
        lines.append(f"{len(lines)+1}. Gained confidence in implementing {module} concepts in real-world scenarios.")

    return "\n".join(lines)


# ─────────────────────────────────────────────────────────────────
# MAIN ENRICHMENT
# ─────────────────────────────────────────────────────────────────

def enrich_excel(path: str):
    wb = openpyxl.load_workbook(path)
    ws = wb.active

    data_rows = [row for row in ws.iter_rows(min_row=5) if row[0].value is not None]

    for idx, row in enumerate(data_rows, start=1):
        sl_cell    = row[0]   # col A: Sl. No.
        date_cell  = row[1]   # col B: Date
        hours_cell = row[2]   # col C: Hours
        mod_cell   = row[3]   # col D: Module
        syl_cell   = row[4]   # col E: Syllabus
        prac_cell  = row[5]   # col F: Practical
        ws_cell    = row[6]   # col G: Work Summary (to write)
        lo_cell    = row[7]   # col H: Learning Outcomes (to write)

        # Fill missing module/syllabus/practical from EXTRA_TOPICS
        if not mod_cell.value and idx in EXTRA_TOPICS:
            t = EXTRA_TOPICS[idx]
            mod_cell.value  = t["module"]
            syl_cell.value  = t["syllabus"]
            prac_cell.value = t["practical"]

        module   = str(mod_cell.value  or "").strip()
        syllabus = str(syl_cell.value  or "").strip()
        practical= str(prac_cell.value or "").strip()

        if not module:
            print(f"  ⚠️  Entry {idx} ({date_cell.value}) has no module — skipping content generation")
            continue

        # Only overwrite col G if it currently holds the raw "Work Summary (Minimum...)" label
        # or is empty. Don't overwrite already-clean content.
        cur_ws = str(ws_cell.value or "").strip()
        if not cur_ws or cur_ws.startswith("Work Summary"):
            ws_cell.value = generate_work_summary(module, syllabus, practical)

        cur_lo = str(lo_cell.value or "").strip()
        if not cur_lo or cur_lo.startswith("Learning Outcomes"):
            lo_cell.value = generate_learnings(module, syllabus, practical)

        date_str = date_cell.value.strftime("%d %b %Y") if isinstance(date_cell.value, datetime) else str(date_cell.value)
        print(f"  ✅ Entry {idx:>3} ({date_str}) — {module}")

    wb.save(path)
    print(f"\n✨ Excel enriched and saved: {path}")


if __name__ == "__main__":
    enrich_excel(EXCEL_FILE)


# AI Presentation Generator Agent

> **Author**: Aditya Rana  
> **Tech Stack**: Python 3.10+ | Groq API (`llama-3.3-70b-versatile`) | `python-pptx` | LibreOffice CLI / Win32 COM  

---

## Overview

An autonomous, multi-step AI Presentation Generator Agent built in Python. Given a PowerPoint template, a topic, a target slide count, and an optional style tone, the agent:

1. **Researches the topic** and constructs a structured executive slide outline.
2. **Generates detailed per-slide content** (action headers, section tags, 3–5 bullet points, key takeaways, and presenter speaker notes).
3. **Populates the template** by programmatically constructing slides off the master layout while enforcing brand typography (Calibri / Calibri Light) and curated color themes (`#27C4B8` Teal, `#0B4F6C` Corporate Navy).
4. **Exports both `.pptx` and `.pdf` files** using automated headless conversion pipelines.

---

## ⚡ Quick Start

### 1. Prerequisites & Setup

Ensure Python 3.10+ is installed on your system.

```bash
# Clone the repository
git clone https://github.com/adityarana121/Presentation-Maker-agent.git
cd "Presentation-Maker-agent"

# Create a virtual environment
python -m venv venv

# Activate the virtual environment
# On Windows (PowerShell):
.\venv\Scripts\Activate.ps1
# On Windows (Command Prompt):
venv\Scripts\activate.bat
# On macOS / Linux:
source venv/bin/activate

# Install required dependencies
pip install -r requirements.txt
```

Configure your environment variables in `.env` (or copy from `.env.example`):

```env
GROQ_API_KEY=your_groq_api_key_here
GROQ_MODEL=llama-3.3-70b-versatile
DEFAULT_NUM_SLIDES=8
DEFAULT_STYLE=professional
DEFAULT_TEMPLATE=TECMOTIV_SOLUTIONS.pptx
OUTPUT_DIR=output
```

---

### 2. Running the Agent

#### Option A: JSON Configuration File (Recommended)
Edit or supply an `input.json` file:
```json
{
  "template": "TECMOTIV_SOLUTIONS.pptx",
  "topic": "Quantum Computing",
  "num_slides": 8,
  "style": "professional"
}
```

Run the pipeline:
```bash
python main.py --input input.json
```

#### Option B: Direct CLI Flags
```bash
python main.py --topic "Quantum Computing" --num_slides 8 --template "TECMOTIV_SOLUTIONS.pptx" --style "professional"
```

#### Output Artifacts
Generated files are written to the `output/` directory:
- `output/Quantum_Computing_presentation.pptx`
- `output/Quantum_Computing_presentation.pdf`

---

## 🏗️ Architecture & Pipeline Design

The system follows a 4-step modular architecture managed by `PresentationAgent` (`src/orchestrator.py`):

```
                                 [ USER INPUT ]
                     (topic, num_slides, template, style)
                                       │
                                       ▼
 ┌─────────────────────────────────────────────────────────────────────────┐
 │ STEP 1: Research & Outline Generation (src/generator/researcher.py)     │
 │ • Analyzes domain context and constructs high-level deck flow          │
 │ • Enforces executive narrative: Title -> Context -> Core -> Outlook    │
 └─────────────────────────────────────┬───────────────────────────────────┘
                                       │
                                       ▼
 ┌─────────────────────────────────────────────────────────────────────────┐
 │ STEP 2: Content Generation (src/generator/content_gen.py)               │
 │ • Generates titles, uppercase category tags, 3-5 bullets, takeaways    │
 │ • Generates 2-4 sentences of presenter speaker notes per slide          │
 └─────────────────────────────────────┬───────────────────────────────────┘
                                       │
                                       ▼
 ┌─────────────────────────────────────────────────────────────────────────┐
 │ STEP 3: Template Population & Layout Engine (src/presenter/builder.py)  │
 │ • Clones template & instantiates slides off Slide Master                │
 │ • Renders styled rounded card containers and key takeaway banners       │
 │ • Enforces explicit left-alignment and safe vertical margins            │
 └─────────────────────────────────────┬───────────────────────────────────┘
                                       │
                                       ▼
 ┌─────────────────────────────────────────────────────────────────────────┐
 │ STEP 4: Presentation Export & PDF Conversion (src/exporter/pdf_exporter.py)
 │ • Saves output .pptx                                                    │
 │ • Converts to .pdf via LibreOffice CLI with Win32 COM fallback          │
 └─────────────────────────────────────────────────────────────────────────┘
```

### Architectural Rationale
1. **Decoupled Modules**: Separating content drafting from slide layout rendering allows swapping LLM backends or visual templates independently without touching core business logic.
2. **Type Safety & Validation**: All intermediate data structures (`PresentationOutline` and `PresentationContent`) are strictly validated using Pydantic schemas.
3. **Token & Latency Efficiency**: Planning the slide outline first ensures cohesive narrative structure before generating individual body text, reducing overall LLM latency to ~10 seconds.

---

## 🎨 Template Analysis & Layout Engine Strategy

### Template Inspection & Layout Challenges
Inspecting the XML layout of `TECMOTIV_SOLUTIONS.pptx` revealed:
- **Blank Canvas**: Slide 1 in the template contains zero shape placeholders.
- **Slide Master Elements**: Branding elements (bottom footer bar, confidential notices, slide numbers, top-right logo `image1.png`, and theme colors `#27C4B8` / `#0B4F6C`) live inside the **Slide Master**.

### Layout Solution
Instead of searching for non-existent body placeholders, `PresentationBuilder`:
1. Clones the template and calls `prs.slides.add_slide(prs.slide_layouts[0])`.
2. Automatically inherits master background graphics, logo, and footer bar on every slide.
3. Calculates safe vertical boundaries (`Top: 0.6"` to `Height: 4.4"`) to avoid overlapping header logos or footer bars.
4. Enforces explicit left paragraph alignment (`PP_ALIGN.LEFT`) across all text frames.
5. Formats Slide 1 as an executive title cover card with 3 clean metadata entries (`Presented by`, `Date`, `Strategic Objective`).

---

## 🧠 AI Tool Usage & Prompt Engineering

### 1. Verbatim Production Prompts

#### A. Step 1: Research & Outline Generation Prompts
**System Prompt:**
```text
You write presentation outlines for executive tech and business presentations. Focus on logical progression, strict topic partitioning, concrete technical topics, and real industry applications. Ensure every slide covers distinct, non-overlapping material so facts are never repeated across slides.
```

**User Prompt:**
```text
Create a structured {num_slides}-slide presentation outline on: "{topic}".

Parameters:
- Slide Count: {num_slides} slides (return exactly {num_slides} items in 'slides')
- Style/Tone: {style}
- Audience: Technical leads, enterprise CTOs, and engineering management

Slide Sequence Structure:
- Slide 1: Title slide introducing the core topic, date, and strategic scope.
- Slide 2: Industry baseline, classical limitations, and fundamental physical challenges (e.g., environmental decoherence & noise).
- Slide 3: Core hardware architectures & qubit implementations (e.g., superconducting circuits vs. trapped ion modalities).
- Slide 4: Quantum algorithms & computational speedups (e.g., Shor's factoring & Grover's database search algorithms).
- Slide 5: Fault-tolerant quantum computing & surface code topologies (focus on physical-to-logical qubit overhead and error thresholds).
- Slide 6: Practical enterprise applications & industry use cases (e.g., post-quantum cryptography, financial portfolio optimization, material simulation).
- Slide 7: Operational benchmarks, quantum volume metrics, and hardware roadmap timeline.
- Slide 8: Strategic recommendations, governance, and immediate enterprise action plan.

CRITICAL: Ensure every slide title and key topic set is completely unique. Do NOT repeat subtopics (e.g., keep error correction code details strictly in Slide 5, algorithm math in Slide 4, and enterprise applications in Slide 6).
```

**Structure & Rationale**:
- **Role Persona**: Sets an analytical, high-level engineering executive persona to avoid low-level generic filler text.
- **Explicit Sequence Mapping**: Enforces a logical narrative flow (Title -> Baseline -> Hardware -> Algorithms -> Fault Tolerance -> Applications -> Benchmarks -> Roadmap) so the deck builds momentum predictably.
- **Topic Partitioning**: Explicitly assigns distinct domains per slide index to prevent information spillover.

---

#### B. Step 2: Per-Slide Content & Speaker Notes Generation Prompts
**System Prompt:**
```text
You are a senior enterprise strategist and lead quantum/AI engineer. You write high-impact, analytical slide content for enterprise executive reviews. Every bullet point must contain concrete technical mechanisms, real industry benchmarks, named algorithms, or specific architecture details. NEVER write circular tautologies or restate the header words in the description. CRITICAL: Maintain strict cross-slide information diversity. Never repeat facts, algorithms, error correction codes, or challenge descriptions that were introduced in earlier slides.
```

**User Prompt:**
```text
Generate detailed slide content for an executive presentation on "{outline.topic}".

Presentation Context:
- Date: {current_date_str}
- Audience: Enterprise CTOs, technical directors, and engineering leads
- Tone: {style}

Slide Outline to Draft:
{self._format_outline(outline)}

Strict Content Requirements per Slide:
1. 'title': Direct, executive slide title (MAX 6 to 8 words / 45 characters max, e.g. "Shor's & Grover's Quantum Algorithms").
2. 'subtitle_tag': A UNIQUE section category header in UPPERCASE (e.g. 'EXECUTIVE OVERVIEW', 'HARDWARE ARCHITECTURE', 'QUANTUM ALGORITHMS', 'FAULT TOLERANCE', 'BENCHMARKS & METRICS', 'ENTERPRISE APPLICATIONS', 'STRATEGIC ROADMAP'). Do NOT repeat category tags across slides.
3. 'bullet_points': 
   - For Slide 1 (Title Slide): EXACTLY 3 metadata lines ("Presented by: Tecmotiv Strategic Advisory", "Date: {current_date_str}", "Scope: {outline.topic} Technology & Strategy Overview").
   - For Slides 2 through {len(outline.slides)} (Content Slides): EXACTLY 4 bullet points per slide. EVERY bullet point MUST follow the structured format: '**Lead-In Header:** Specific technical explanation...'
   - The Lead-In Header must be a bold 2-4 word topic label (e.g. '**Qubit Implementation:**', '**Decoherence Mitigation:**', '**Shor\'s Algorithm:**').
   - The explanation MUST state concrete facts, real hardware/software examples (e.g., IBM Eagle, trapped ions, surface codes, RSA impact), metrics, or operational mechanisms.
   - DO NOT write circular definitions (e.g. DO NOT say "Quantum Noise: Noise in quantum systems.").
   - ANTI-REPETITION RULE: Every slide MUST cover completely distinct information. If Slide 2 discusses environmental noise and decoherence, Slide 5 must focus purely on surface code lattice topologies, physical-to-logical qubit ratios, and threshold rates (10^-3) WITHOUT re-explaining noise basics or repeating Shor/Steane error correction code mentions from earlier slides.
4. 'key_takeaway': A single sharp, analytical insight summarizing the key technical or business conclusion of the slide.
5. 'speaker_notes': 2 to 3 concise sentences providing deep technical context for the presenter.
```

**Structure & Rationale**:
- **Date Grounding**: Injects exact current timestamp (`July 2026`) into the prompt to ensure realistic dates and eliminate generic placeholder brackets like `[Date]` or `[Name]`.
- **Enforced Lead-In Pattern**: Enforces `**Lead-In Header:** Explanation` syntax for visual emphasis in PowerPoint.
- **Negative Constraints**: Strictly forbids tautological definitions and cross-slide repetition.

---

### 2. Prompt Iteration & Refinement History

#### Iteration 1: Fixing Paragraph Drift & Formatting Mechanical Inconsistencies
- **Initial Prompt**:
  > *"Write 3-4 bullet points for each slide. Include a title, category tag, and summary takeaway."*
- **Issue Encountered**:
  The LLM returned full 30-word paragraphs instead of bullet points. In PowerPoint, these rendered as massive walls of text that overflowed the visual card background container.
- **Iteration & Fix**:
  Added explicit structural formatting requirements: `EVERY bullet point MUST follow the structured format: '**Lead-In Header:** Specific technical explanation...'` and capped titles at 6-8 words / 45 characters. This transformed unstructured text blocks into scannable, executive-ready callouts with bold primary anchors.

#### Iteration 2: Eliminating Cross-Slide Content Duplication & Repetition
- **Initial Prompt**:
  > *"Generate content for each slide based on the provided outline topics."*
- **Issue Encountered**:
  Slide 2 ("Problem Statement") discussed Shor and Steane error correction codes. Slide 5 ("Quantum Error Correction") repeated almost the exact same sentences about Shor/Steane codes being essential for reliable computing.
- **Iteration & Fix**:
  Updated both the System Prompt and Content Prompt with explicit **ANTI-REPETITION RULES** and context awareness:
  > *"CRITICAL: Maintain strict cross-slide information diversity. If Slide 2 covers noise & decoherence basics, Slide 5 must focus deeply on physical qubit overhead, surface codes, and logical qubit thresholds without re-explaining what noise is or repeating Shor/Steane mentions."*

---

### 3. Verbatim Debugging Prompt (python-pptx Null-Check Issue)

During development of Step 3 (`src/presenter/builder.py`), attaching speaker notes threw a runtime `AttributeError` on certain slides because `slide.notes_slide` returned `None` or its `notes_text_frame` was uninitialized in cloned template instances.

**Actual Debugging Prompt Sent to LLM**:
```text
I am using python-pptx to populate slides from an existing PowerPoint template (TECMOTIV_SOLUTIONS.pptx). 
When I attempt to write speaker notes using `slide.notes_slide.notes_text_frame.text = slide_data.speaker_notes`, I occasionally get an AttributeError: 'NoneType' object has no attribute 'notes_text_frame' or notes_slide is not initialized automatically for slides instantiated from blank master layouts.

Here is my current code snippet:
```python
notes_slide = slide.notes_slide
notes_slide.notes_text_frame.text = slide_data.speaker_notes
```

How do I safely initialize and access `notes_slide` in python-pptx to set text without throwing NoneType exceptions across all PowerPoint versions?
```

**Resolution Applied**:
Wrapped the speaker notes assignment in a defensive check checking both `slide.notes_slide` and `notes_slide.notes_text_frame`:
```python
if slide_data.speaker_notes:
    try:
        notes_slide = slide.notes_slide
        if notes_slide and notes_slide.notes_text_frame:
            notes_slide.notes_text_frame.text = slide_data.speaker_notes
    except Exception as e:
        logger.warning(f"Could not attach speaker notes to slide {index + 1}: {e}")
```

---

## 📄 AI Disclosure Statement

In accordance with project transparency guidelines:

- **LLM (`llama-3.3-70b-versatile` via Groq)**: Topic research, outline structure generation, slide copywriting, and presenter speaker notes.
- **`python-pptx`**: Template cloning, slide master instantiation, card layout rendering, visual formatting, and speaker notes injection.
- **LibreOffice CLI & Win32 COM**: Automated headless PPTX to PDF conversion.
- **LLM Debugging**: Used LLM during development to resolve `python-pptx` text frame initialization null checks when writing slide speaker notes.

---

## 🛡️ Resilience & Production Engineering

| Feature | Design Implementation |
| :--- | :--- |
| **API Retry Mechanism** | Exponential backoff (up to 3 retries with 1.5s multiplier) for rate limits and transient network failures. |
| **Schema Validation** | Strips markdown blocks (` ```json `), parses JSON, and validates models via Pydantic. |
| **File Lock Fallback** | Catches `PermissionError` if `.pptx` is open in PowerPoint UI and saves to a clean fallback output path. |
| **PDF Conversion Fallback** | Prefers headless LibreOffice CLI (`soffice`), with native Win32 COM PowerPoint automation fallback. |
| **Topic Enrichment** | Automatically expands short 1-word topics into structured presentation themes. |

---

## ⚠️ Limitations & Future Roadmap

### Limitations
1. **Text Container Rendering**: Renders structured text card containers and takeaway banners; does not auto-generate native PowerPoint statistical charts yet.
2. **Aspect Ratio**: Calibrated for standard 16:9 widescreen presentations (10.0" x 5.625").

### Future Enhancements
1. **Single-Slide Targeted Edits**: Add CLI flag `--regen-slide 3` to modify individual slides without regenerating the full presentation.
2. **Native Chart Engine**: Integrate `pptx.chart.data.CategoryChartData` to convert statistical data into native PowerPoint charts.
3. **Dynamic Visual Grids**: Support multi-column comparison grids and vertical 3-pillar card layouts based on slide content type.

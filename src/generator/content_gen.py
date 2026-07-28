"""
Per-slide presentation content and speaker notes generator.
"""

from datetime import datetime
import logging
import re
from src.llm_client import LLMClient
from src.models import PresentationOutline, PresentationContent

logger = logging.getLogger("ContentGenerator")


def is_tautological_bullet(bullet: str) -> bool:
    """Detects circular definitions where the description starts by repeating the lead-in header."""
    if ":" not in bullet:
        return False
    
    parts = bullet.split(":", 1)
    lead_in = parts[0].replace("*", "").strip()
    desc = parts[1].replace("*", "").strip()

    stopwords = {"and", "or", "the", "in", "of", "to", "for", "a", "an", "is", "are", "with", "on", "by", "as", "such", "that"}
    lead_words = [w.lower() for w in re.findall(r'\b[a-zA-Z]{3,}\b', lead_in) if w.lower() not in stopwords]
    
    if not lead_words:
        return False

    # Check the first 4 words of description
    first_four_words = [w.lower() for w in re.findall(r'\b[a-zA-Z]{3,}\b', desc)[:4]]
    if not first_four_words:
        return False

    # If 2 or more lead words appear in the FIRST 4 words of description (e.g., "Performance Metrics: Performance metrics are..."), it's circular!
    start_overlap = sum(1 for w in first_four_words if w in lead_words)
    if start_overlap >= 2 or (len(lead_words) == 1 and first_four_words[0] in lead_words):
        return True

    # Also check if description starts with "X is a/an/the Y..."
    first_six_str = " ".join([w.lower() for w in re.findall(r'\b[a-zA-Z]{3,}\b', desc)[:6]])
    lead_in_str = " ".join(lead_words)
    if lead_in_str in first_six_str:
        return True

    return False


class ContentGenerator:
    """Generates detailed slide copy, bullet points, and speaker notes from an outline."""

    def __init__(self, llm_client: LLMClient):
        self.llm = llm_client

    def generate_presentation_content(
        self, outline: PresentationOutline, style: str = "professional"
    ) -> PresentationContent:
        current_date_str = datetime.now().strftime("%B %Y")
        logger.info(f"Drafting content for {len(outline.slides)} slides on '{outline.topic}'...")

        system_prompt = (
            "You are a senior enterprise strategist and lead quantum/AI engineer. "
            "You write high-impact, highly analytical slide content for enterprise executive reviews. "
            "Every bullet point MUST state concrete technical mechanisms, named commercial systems (e.g., IBM Eagle, Sycamore, D-Wave), "
            "exact algorithms (Shor's, Grover's), specific standards (NIST FIPS 203 ML-KEM, IEEE P2881, QED-C), or numerical metrics. "
            "NEVER write circular tautologies or restate the header words in the description. "
            "CRITICAL: Maintain strict cross-slide information diversity. Never repeat facts, algorithms, or error correction codes introduced in earlier slides."
        )

        prompt = f"""
Generate detailed slide content for an executive presentation on "{outline.topic}".

Presentation Context:
- Date: {current_date_str}
- Audience: Enterprise CTOs, technical directors, and engineering leads
- Tone: {style}

Slide Outline to Draft:
{self._format_outline(outline)}

FEW-SHOT EXAMPLES (STUDY THESE CAREFULLY):

[BAD - FORBIDDEN CIRCULAR BULLETS]:
- ❌ "**Performance Metrics:** Performance metrics measure the performance of a quantum computer using performance benchmarks."
- ❌ "**Exponential Scaling Challenges:** Classical systems face exponential scaling challenges due to exponential complexity."
- ❌ "**Immediate Action Plan:** An immediate action plan is a plan for short-term development and deployment."

[GOOD - REQUIRED ANALYTICAL BULLETS]:
- ✅ "**IBM Quantum System Two:** Integrates 133-qubit Heron processors achieving 5,000+ Quantum Volume and 2x error reduction over Eagle."
- ✅ "**RSA-2048 Security Impact:** Shor's algorithm reduces 2048-bit prime factorization from 10,000 classical years to <8 hours on 4,000 logical qubits."
- ✅ "**NIST FIPS 203 Transition:** Mandates enterprise migration of TLS 1.3 key exchange endpoints to ML-KEM lattice cryptography by Q4 2026."
- ✅ "**Surface Code Overhead:** Requires 1,000 physical qubits per logical qubit at a physical error threshold rate below 10^-3."

Strict Content Requirements per Slide:
1. 'title': Direct, executive slide title (MAX 6 to 8 words / 45 characters max, e.g. "Shor's & Grover's Quantum Algorithms").
2. 'subtitle_tag': A UNIQUE section category header in UPPERCASE (e.g. 'EXECUTIVE OVERVIEW', 'HARDWARE ARCHITECTURE', 'QUANTUM ALGORITHMS', 'FAULT TOLERANCE', 'BENCHMARKS & METRICS', 'ENTERPRISE APPLICATIONS', 'STRATEGIC ROADMAP'). Do NOT repeat category tags across slides.
3. 'bullet_points': 
   - For Slide 1 (Title Slide): EXACTLY 3 metadata lines ("Presented by: Tecmotiv Strategic Advisory", "Date: {current_date_str}", "Scope: {outline.topic} Technology & Strategy Overview").
   - For Slides 2 through {len(outline.slides)} (Content Slides): EXACTLY 4 bullet points per slide. EVERY bullet point MUST follow the structured format: '**Lead-In Header:** Specific technical explanation...'
   - The Lead-In Header must be a bold 2-4 word topic label (e.g. '**Superconducting Modality:**', '**Decoherence Mitigation:**', '**Shor\'s Algorithm:**').
   - The explanation MUST state concrete facts, real hardware/software examples (e.g., IBM Eagle 127-qubit, Sycamore, D-Wave Advantage, surface code lattices, RSA-2048 impact), verified standards bodies (e.g., QED-C, NIST FIPS 203 ML-KEM, IEEE P2881), metrics, or operational mechanisms.
   - DO NOT write circular definitions (e.g. DO NOT say "Quantum Noise: Noise in quantum systems.").
   - ANTI-REPETITION RULE: Every slide MUST cover completely distinct information. If Slide 2 discusses environmental noise and decoherence, Slide 5 must focus purely on surface code lattice topologies, physical-to-logical qubit ratios, and threshold rates (10^-3) WITHOUT re-explaining noise basics or repeating Shor/Steane error correction code mentions from earlier slides.
4. 'key_takeaway': A single sharp, analytical insight summarizing the key technical or business conclusion of the slide.
5. 'speaker_notes': 2 to 3 concise sentences providing deep technical context for the presenter.
"""

        for attempt in range(1, 4):
            try:
                content = self.llm.generate_json(
                    prompt=prompt,
                    system_prompt=system_prompt,
                    response_model=PresentationContent
                )

                # Check for <4 bullets AND tautological bullets
                invalid_bullet_slides = []
                tautological_slides = []

                for s in content.slides:
                    if s.slide_number > 1:
                        if len(s.bullet_points) < 4:
                            invalid_bullet_slides.append(s.slide_number)
                        for bullet in s.bullet_points:
                            if is_tautological_bullet(bullet):
                                tautological_slides.append((s.slide_number, bullet))

                if not invalid_bullet_slides and not tautological_slides:
                    logger.info(f"Generated slide copy for {len(content.slides)} slides (all slides passed 4-bullet and non-circular verification).")
                    return content
                
                if invalid_bullet_slides:
                    logger.warning(f"Attempt {attempt}/3: Found slide(s) {invalid_bullet_slides} with <4 bullets.")
                if tautological_slides:
                    for s_num, b_text in tautological_slides:
                        logger.warning(f"Attempt {attempt}/3: Slide {s_num} contains circular bullet: '{b_text}'")
                
                logger.warning(f"Retrying content generation (Attempt {attempt}/3)...")
            except Exception as e:
                logger.warning(f"Attempt {attempt}/3 content generation failed: {e}")
                if attempt == 3:
                    raise e

        logger.info(f"Generated slide copy for {len(content.slides)} slides.")
        return content

    @staticmethod
    def _format_outline(outline: PresentationOutline) -> str:
        items = []
        for slide in outline.slides:
            items.append(
                f"Slide {slide.slide_number} ({slide.slide_type}): '{slide.title}' -> Topics: {', '.join(slide.key_topics)}"
            )
        return "\n".join(items)

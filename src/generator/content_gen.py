"""
Per-slide presentation content and speaker notes generator.
"""

from datetime import datetime
import logging
from src.llm_client import LLMClient
from src.models import PresentationOutline, PresentationContent

logger = logging.getLogger("ContentGenerator")


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
            "You write high-impact, analytical slide content for enterprise executive reviews. "
            "Every bullet point must contain concrete technical mechanisms, real industry benchmarks, named algorithms, or specific architecture details. "
            "NEVER write circular tautologies or restate the header words in the description. "
            "CRITICAL: Maintain strict cross-slide information diversity. Never repeat facts, algorithms, error correction codes, or challenge descriptions that were introduced in earlier slides."
        )

        prompt = f"""
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
   - The explanation MUST state concrete facts, real hardware/software examples (e.g., IBM Eagle 127-qubit, trapped ion architectures, surface code lattices, RSA-2048 impact), verified standards bodies (e.g., QED-C, NIST FIPS 203 ML-KEM, IEEE P2881), metrics, or operational mechanisms.
   - DO NOT hallucinate organization names (e.g., use 'QED-C' or 'IEEE Quantum Initiative', NEVER invent fake consortium names).
   - For Governance and Roadmap slides, cite specific frameworks such as NIST Cybersecurity Framework, ISO/IEC 27001 risk controls, and PQC migration timelines.
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

                # Check bullet counts for content slides (slides 2+)
                invalid_slides = [s for s in content.slides if s.slide_number > 1 and len(s.bullet_points) < 4]
                if not invalid_slides:
                    logger.info(f"Generated slide copy for {len(content.slides)} slides (all content slides verified with 4 bullets).")
                    return content
                
                logger.warning(
                    f"Attempt {attempt}/3: Found {len(invalid_slides)} slide(s) with <4 bullets. Retrying JSON generation..."
                )
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

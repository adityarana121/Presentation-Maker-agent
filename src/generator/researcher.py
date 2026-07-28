"""
Topic research and presentation outline generator.
"""

import logging
from src.llm_client import LLMClient
from src.models import PresentationOutline

logger = logging.getLogger("Researcher")


class TopicResearcher:
    """Generates structured presentation outlines based on topic research."""

    def __init__(self, llm_client: LLMClient):
        self.llm = llm_client

    def generate_outline(self, topic: str, num_slides: int = 8, style: str = "professional") -> PresentationOutline:
        logger.info(f"Generating slide outline for '{topic}' ({num_slides} slides)...")

        system_prompt = (
            "You write presentation outlines for executive tech and business presentations. "
            "Focus on logical progression, strict topic partitioning, concrete technical topics, and real industry applications. "
            "Ensure every slide covers distinct, non-overlapping material so facts are never repeated across slides."
        )

        prompt = f"""
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
"""

        outline = self.llm.generate_json(
            prompt=prompt,
            system_prompt=system_prompt,
            response_model=PresentationOutline
        )

        # Enforce exact slide count indexing
        if len(outline.slides) != num_slides:
            logger.warning(f"Adjusting generated slide count ({len(outline.slides)}) to target ({num_slides}).")
            for idx, slide in enumerate(outline.slides, start=1):
                slide.slide_number = idx

        logger.info(f"Outline generated with {len(outline.slides)} slides.")
        return outline

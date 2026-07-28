"""
Step 1: Research & Outline Generation Module.
Researches topic and generates a structured slide outline using LLM.
"""

import logging
from src.llm_client import LLMClient
from src.models import PresentationOutline

logger = logging.getLogger("Researcher")


class TopicResearcher:
    """Researches presentation topic and generates structured slide outline."""

    def __init__(self, llm_client: LLMClient):
        self.llm = llm_client

    def generate_outline(self, topic: str, num_slides: int = 8, style: str = "professional") -> PresentationOutline:
        """
        Generates a structured presentation outline for the given topic and target slide count.
        """
        logger.info(f"Starting Step 1: Researching topic '{topic}' for {num_slides} slides (Style: {style})...")

        system_prompt = (
            "You are an expert executive presentation strategist and researcher at a top management consulting firm. "
            "Your task is to structure a compelling, logically flowing slide outline for a high-impact corporate deck."
        )

        prompt = f"""
Research and create a structured {num_slides}-slide presentation outline on the topic: "{topic}".

Parameters:
- Target Slide Count: {num_slides} slides (MUST produce exactly {num_slides} slides in the 'slides' list)
- Presentation Tone/Style: {style}
- Audience: Enterprise decision-makers, tech leads, and strategic partners

Requirements for Slide Flow:
1. Slide 1 MUST be a 'title' slide (Executive Title & Subtitle Context).
2. Slide 2 should set the context, problem statement, or industry overview.
3. Middle slides ({num_slides - 3} slides) should delve into core pillars, technical architecture, strategic benefits, and real-world use cases.
4. The final slide should be a 'summary' slide (Key Takeaways, Strategic Recommendations, & Next Steps).

Ensure each slide has a clear, professional title and 3-5 specific topics to cover.
"""

        outline = self.llm.generate_json(
            prompt=prompt,
            system_prompt=system_prompt,
            response_model=PresentationOutline
        )

        # Enforce exact slide count constraint
        if len(outline.slides) != num_slides:
            logger.warning(
                f"LLM generated {len(outline.slides)} slides instead of target {num_slides}. Re-indexing slide numbers."
            )
            for idx, slide in enumerate(outline.slides, start=1):
                slide.slide_number = idx

        logger.info(f"Step 1 Complete: Generated outline with {len(outline.slides)} slides.")
        return outline

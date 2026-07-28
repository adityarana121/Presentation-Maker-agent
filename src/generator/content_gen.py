"""
Step 2: Content Generation Module.
Generates per-slide detailed content (title, 3-5 bullet points, takeaway, speaker notes)
based on the outline from Step 1.
"""

import logging
from src.llm_client import LLMClient
from src.models import PresentationOutline, PresentationContent, SlideContent

logger = logging.getLogger("ContentGenerator")


class ContentGenerator:
    """Generates detailed slide-by-slide content from outline."""

    def __init__(self, llm_client: LLMClient):
        self.llm = llm_client

    def generate_presentation_content(
        self, outline: PresentationOutline, style: str = "professional"
    ) -> PresentationContent:
        """
        Generates per-slide content for all slides in the outline.
        """
        logger.info(f"Starting Step 2: Generating detailed slide content for '{outline.topic}' ({len(outline.slides)} slides)...")

        system_prompt = (
            "You are a principal management consultant and slide presentation copywriter. "
            "Your writing is concise, analytical, highly structured, and impactful. "
            "Avoid generic filler words; use precise corporate and technical terminology."
        )

        prompt = f"""
Generate detailed slide content for a {len(outline.slides)}-slide presentation on "{outline.topic}".

Presentation Context & Outline:
- Tone/Style: {style}
- Target Audience: {outline.target_audience}
- Core Narrative: {outline.core_narrative}

Slides to Generate:
{self._format_outline_for_prompt(outline)}

Strict Content Requirements per Slide:
1. 'title': Action-oriented header (e.g., "Quantum Advantage in Financial Cryptography").
2. 'subtitle_tag': Section category tag in UPPERCASE (e.g., "CORE ARCHITECTURE", "MARKET IMPACT").
3. 'bullet_points': EXACTLY 3 to 5 clear, high-impact bullet points:
   - Each bullet point must be 10-15 words max.
   - Start with action verbs or bold key terms.
   - For Slide 1 (Title Slide), bullet points should represent executive presentation metadata (e.g. Presenter, Date, Strategic Objective).
4. 'key_takeaway': A concise single-sentence strategic insight summarizing the slide.
5. 'speaker_notes': 2 to 4 detailed sentences providing contextual background commentary for the presenter.

Ensure tone, depth, and vocabulary remain strictly consistent across all slides.
"""

        content = self.llm.generate_json(
            prompt=prompt,
            system_prompt=system_prompt,
            response_model=PresentationContent
        )

        logger.info(f"Step 2 Complete: Generated detailed content for {len(content.slides)} slides.")
        return content

    @staticmethod
    def _format_outline_for_prompt(outline: PresentationOutline) -> str:
        formatted = []
        for slide in outline.slides:
            formatted.append(
                f"- Slide {slide.slide_number} ({slide.slide_type}): Title='{slide.title}', "
                f"Topics={', '.join(slide.key_topics)}"
            )
        return "\n".join(formatted)

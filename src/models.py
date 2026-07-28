"""
Pydantic Data Models for Slide Generation Pipeline.
Defines schemas for outlines, detailed slide content, and pipeline input/output.
"""

from typing import List, Optional
from pydantic import BaseModel, Field


class InputRequest(BaseModel):
    """Input parameters provided by user or JSON input file."""
    template: str = Field(default="TECMOTIV_SOLUTIONS.pptx", description="Path to input template presentation")
    topic: str = Field(..., description="Topic of the presentation")
    num_slides: int = Field(default=8, ge=3, le=20, description="Target number of slides (3-20)")
    style: str = Field(default="professional", description="Presentation style tone (e.g. professional, executive, technical)")


class SlideOutlineItem(BaseModel):
    """Outline structure for a single slide generated in Step 1."""
    slide_number: int = Field(..., description="1-based slide index")
    title: str = Field(..., description="Concise, impactful title for the slide")
    key_topics: List[str] = Field(..., description="High-level topics/points to cover in this slide")
    slide_type: str = Field(
        default="content",
        description="Type of slide layout: 'title', 'overview', 'content', 'comparison', or 'summary'"
    )


class PresentationOutline(BaseModel):
    """Complete presentation outline generated in Step 1."""
    topic: str
    num_slides: int
    target_audience: str = Field(..., description="Identified audience for tone alignment")
    core_narrative: str = Field(..., description="Central theme or thesis statement of presentation")
    slides: List[SlideOutlineItem]


class SlideContent(BaseModel):
    """Detailed generated content for a single slide (Step 2)."""
    slide_number: int
    title: str = Field(..., description="Refined, professional slide header title")
    subtitle_tag: str = Field(..., description="Category tag or section subtitle (e.g., 'STRATEGIC OVERVIEW')")
    bullet_points: List[str] = Field(
        ...,
        min_items=3,
        max_items=5,
        description="3 to 5 clear, actionable bullet points (10-15 words per bullet)"
    )
    key_takeaway: Optional[str] = Field(
        default=None,
        description="Highlighted bottom summary or core callout for this slide"
    )
    speaker_notes: str = Field(
        ...,
        description="Detailed background commentary for the presenter (2-4 sentences)"
    )


class PresentationContent(BaseModel):
    """Complete per-slide content payload generated in Step 2."""
    topic: str
    style: str
    slides: List[SlideContent]


class PipelineOutput(BaseModel):
    """Final output object returned by the orchestration engine."""
    status: str
    topic: str
    num_slides: int
    pptx_path: str
    pdf_path: Optional[str] = None
    pdf_converted: bool = False
    conversion_method: str = "none"
    execution_time_seconds: float = 0.0
    error_message: Optional[str] = None

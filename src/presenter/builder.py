"""
Step 3: Presentation Builder Module.
Populates TECMOTIV_SOLUTIONS.pptx template preserving slide master branding,
applying Tecmotiv typography, card layouts, and speaker notes.
"""

import os
import logging
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

from src.config import (
    DEFAULT_TEMPLATE_PATH,
    COLOR_PRIMARY_NAVY,
    COLOR_ACCENT_TEAL,
    COLOR_VIOLET_PRIMARY,
    COLOR_CARD_BG,
    COLOR_TEXT_MAIN,
    COLOR_TEXT_MUTED,
    COLOR_WHITE,
    FONT_HEADING,
    FONT_BODY,
)
from src.models import PresentationContent, SlideContent

logger = logging.getLogger("PresentationBuilder")


class PresentationBuilder:
    """Builds PPTX presentation populated with generated content based on Tecmotiv template."""

    def __init__(self, template_path: str = DEFAULT_TEMPLATE_PATH):
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"Template file not found at path: {template_path}")
        self.template_path = template_path

    def build_presentation(self, content: PresentationContent, output_pptx_path: str) -> str:
        """
        Clones template and builds all slides off the master layout.
        """
        logger.info(f"Starting Step 3: Populating template '{self.template_path}' with {len(content.slides)} slides...")

        prs = Presentation(self.template_path)
        blank_layout = prs.slide_layouts[0]

        # Delete existing initial blank slides in template to build fresh content set
        while len(prs.slides) > 0:
            rId = prs.slides._sldIdLst[0].rId
            prs.part.drop_rel(rId)
            del prs.slides._sldIdLst[0]

        # Build each slide
        for index, slide_data in enumerate(content.slides):
            slide = prs.slides.add_slide(blank_layout)
            if index == 0:
                self._build_title_slide(slide, slide_data, content.topic)
            else:
                self._build_content_slide(slide, slide_data)

            # Add speaker notes
            if slide_data.speaker_notes:
                notes_slide = slide.notes_slide
                text_frame = notes_slide.notes_text_frame
                text_frame.text = slide_data.speaker_notes

        prs.save(output_pptx_path)
        logger.info(f"Step 3 Complete: Presentation saved to '{output_pptx_path}'")
        return output_pptx_path

    def _build_title_slide(self, slide, data: SlideContent, topic: str):
        """Constructs Slide 1 (Title Slide) with executive presentation hero branding."""
        # Top Category Tag
        tag_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.8), Inches(8.4), Inches(0.4))
        tf_tag = tag_box.text_frame
        tf_tag.word_wrap = True
        p_tag = tf_tag.paragraphs[0]
        p_tag.text = data.subtitle_tag.upper() if data.subtitle_tag else "STRATEGIC PRESENTATION"
        p_tag.font.name = FONT_BODY
        p_tag.font.size = Pt(11)
        p_tag.font.bold = True
        p_tag.font.color.rgb = COLOR_ACCENT_TEAL

        # Main Deck Title
        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.2), Inches(8.4), Inches(1.4))
        tf_title = title_box.text_frame
        tf_title.word_wrap = True
        p_title = tf_title.paragraphs[0]
        p_title.text = data.title
        p_title.font.name = FONT_HEADING
        p_title.font.size = Pt(32)
        p_title.font.bold = True
        p_title.font.color.rgb = COLOR_PRIMARY_NAVY

        # Subtle Accent Horizontal Line
        line = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(2.7), Inches(2.5), Inches(0.04)
        )
        line.fill.solid()
        line.fill.fore_color.rgb = COLOR_ACCENT_TEAL
        line.line.color.rgb = COLOR_ACCENT_TEAL

        # Executive Context Card Container
        card = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(2.9), Inches(8.4), Inches(1.8)
        )
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_CARD_BG
        card.line.color.rgb = COLOR_ACCENT_TEAL
        card.line.width = Pt(1)

        # Content bullets inside Title Card
        tf_card = card.text_frame
        tf_card.word_wrap = True
        tf_card.margin_left = Inches(0.25)
        tf_card.margin_top = Inches(0.2)
        tf_card.margin_right = Inches(0.25)

        for i, bullet in enumerate(data.bullet_points):
            p = tf_card.paragraphs[0] if i == 0 else tf_card.add_paragraph()
            p.text = bullet
            p.font.name = FONT_BODY
            p.font.size = Pt(13)
            p.font.color.rgb = COLOR_TEXT_MAIN
            p.space_after = Pt(6)

    def _build_content_slide(self, slide, data: SlideContent):
        """Constructs Content Slide with category header, content card, and bottom takeaway bar."""
        # Category Tag
        tag_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.45), Inches(8.4), Inches(0.35))
        tf_tag = tag_box.text_frame
        tf_tag.word_wrap = True
        tf_tag.margin_top = Inches(0)
        tf_tag.margin_bottom = Inches(0)
        p_tag = tf_tag.paragraphs[0]
        p_tag.text = data.subtitle_tag.upper()
        p_tag.font.name = FONT_BODY
        p_tag.font.size = Pt(10)
        p_tag.font.bold = True
        p_tag.font.color.rgb = COLOR_ACCENT_TEAL

        # Main Slide Title
        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.75), Inches(8.4), Inches(0.65))
        tf_title = title_box.text_frame
        tf_title.word_wrap = True
        tf_title.margin_top = Inches(0)
        p_title = tf_title.paragraphs[0]
        p_title.text = data.title
        p_title.font.name = FONT_HEADING
        p_title.font.size = Pt(22)
        p_title.font.bold = True
        p_title.font.color.rgb = COLOR_PRIMARY_NAVY

        # Main Content Card Background
        card_top = Inches(1.35)
        card_height = Inches(2.75)
        card = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), card_top, Inches(8.4), card_height
        )
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_CARD_BG
        card.line.color.rgb = COLOR_ACCENT_TEAL
        card.line.width = Pt(0.75)

        # Bullet Points Text inside Card
        tf_card = card.text_frame
        tf_card.word_wrap = True
        tf_card.margin_left = Inches(0.3)
        tf_card.margin_top = Inches(0.2)
        tf_card.margin_right = Inches(0.3)
        tf_card.margin_bottom = Inches(0.2)

        for i, bullet in enumerate(data.bullet_points):
            p = tf_card.paragraphs[0] if i == 0 else tf_card.add_paragraph()
            p.text = f"•  {bullet}"
            p.font.name = FONT_BODY
            p.font.size = Pt(13)
            p.font.color.rgb = COLOR_TEXT_MAIN
            p.space_after = Pt(8)

        # Key Takeaway Banner at bottom of content area
        if data.key_takeaway:
            takeaway_top = Inches(4.2)
            takeaway_box = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, Inches(0.8), takeaway_top, Inches(8.4), Inches(0.6)
            )
            takeaway_box.fill.solid()
            takeaway_box.fill.fore_color.rgb = COLOR_PRIMARY_NAVY
            takeaway_box.line.color.rgb = COLOR_ACCENT_TEAL
            takeaway_box.line.width = Pt(1.5)

            tf_takeaway = takeaway_box.text_frame
            tf_takeaway.word_wrap = True
            tf_takeaway.margin_left = Inches(0.2)
            tf_takeaway.margin_right = Inches(0.2)
            p_takeaway = tf_takeaway.paragraphs[0]
            p_takeaway.text = f"KEY TAKEAWAY: {data.key_takeaway}"
            p_takeaway.font.name = FONT_BODY
            p_takeaway.font.size = Pt(11)
            p_takeaway.font.bold = True
            p_takeaway.font.color.rgb = COLOR_WHITE
            p_takeaway.alignment = PP_ALIGN.LEFT

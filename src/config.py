"""
Configuration and Design System for AI Presentation Generator Agent.
Defines theme colors, typography, layout bounds, and environment variables.
"""

import os
from pathlib import Path
from dotenv import load_dotenv
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

# Load environment variables from .env file
load_dotenv()

# Base Paths
BASE_DIR = Path(__file__).resolve().parent.parent
DEFAULT_TEMPLATE_PATH = os.getenv("DEFAULT_TEMPLATE", str(BASE_DIR / "TECMOTIV_SOLUTIONS.pptx"))
OUTPUT_DIR = Path(os.getenv("OUTPUT_DIR", str(BASE_DIR / "output")))
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# LLM Configuration
GROQ_API_KEY = os.getenv("GROQ_API_KEY")
GROQ_MODEL = os.getenv("GROQ_MODEL", "llama-3.1-8b-instant")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY")

# Tecmotiv Solutions Theme Colors (Derived from Slide Master XML & Brand Logo)
COLOR_PRIMARY_NAVY = RGBColor(0x0B, 0x4F, 0x6C)      # #0B4F6C - Dark Blue / Footer Bar
COLOR_ACCENT_TEAL = RGBColor(0x27, 0xC4, 0xB8)       # #27C4B8 - Tecmotiv Teal (Logo & Accent Line)
COLOR_VIOLET_PRIMARY = RGBColor(0x92, 0x27, 0x8F)    # #92278F - Violet Theme Primary
COLOR_VIOLET_LIGHT = RGBColor(0x9B, 0x57, 0xD3)      # #9B57D3 - Violet Theme Accent
COLOR_CARD_BG = RGBColor(0xF4, 0xF9, 0xF9)           # #F4F9F9 - Light Teal Tint Card Background
COLOR_TEXT_MAIN = RGBColor(0x1F, 0x29, 0x37)         # #1F2937 - Dark Charcoal Body Text
COLOR_TEXT_MUTED = RGBColor(0x6B, 0x72, 0x80)        # #6B7280 - Slate Muted Text
COLOR_WHITE = RGBColor(0xFF, 0xFF, 0xFF)             # Pure White

# Typography Settings (Calibri / Calibri Light per Tecmotiv theme)
FONT_HEADING = "Calibri Light"
FONT_BODY = "Calibri"

# Slide Dimensions & Bounds (10.0" x 5.625" 16:9 Widescreen)
SLIDE_WIDTH_INCHES = 10.0
SLIDE_HEIGHT_INCHES = 5.625

# Safe Content Area (Prevents overlap with top-right logo and bottom Confidential bar)
CONTENT_TOP_INCHES = 0.6
CONTENT_LEFT_INCHES = 0.8
CONTENT_WIDTH_INCHES = 8.4
CONTENT_HEIGHT_MAX_INCHES = 4.4  # Leaves ~0.6" at bottom for Confidential bar

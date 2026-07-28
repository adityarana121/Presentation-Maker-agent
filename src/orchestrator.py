"""
Presentation Pipeline Orchestrator.
Coordinates end-to-end presentation generation across all 4 steps:
Research -> Content Generation -> Template Population -> PPTX/PDF Export.
"""

import time
import logging
from pathlib import Path
from typing import Optional

from src.config import OUTPUT_DIR, DEFAULT_TEMPLATE_PATH
from src.llm_client import LLMClient
from src.models import InputRequest, PipelineOutput
from src.generator.researcher import TopicResearcher
from src.generator.content_gen import ContentGenerator
from src.presenter.builder import PresentationBuilder
from src.exporter.pdf_exporter import PDFExporter

logger = logging.getLogger("PresentationOrchestrator")


class PresentationAgent:
    """Orchestrator managing the multi-step AI presentation generation pipeline."""

    def __init__(self, llm_provider: str = "auto"):
        self.llm_client = LLMClient(provider=llm_provider)
        self.researcher = TopicResearcher(self.llm_client)
        self.content_gen = ContentGenerator(self.llm_client)
        self.pdf_exporter = PDFExporter()

    def run(self, input_request: InputRequest) -> PipelineOutput:
        """Executes full pipeline for the given input request."""
        start_time = time.time()
        logger.info("==========================================================================")
        logger.info(f"STARTING PRESENTATION GENERATION PIPELINE FOR TOPIC: '{input_request.topic}'")
        logger.info("==========================================================================")

        try:
            # 1. Sanitize Topic & File Paths
            clean_topic = "".join(c if c.isalnum() else "_" for c in input_request.topic).strip("_")
            clean_topic = "_".join(filter(None, clean_topic.split("_")))
            
            output_pptx_filename = f"{clean_topic}_presentation.pptx"
            output_pdf_filename = f"{clean_topic}_presentation.pdf"
            
            pptx_output_path = str(OUTPUT_DIR / output_pptx_filename)
            pdf_output_path = str(OUTPUT_DIR / output_pdf_filename)

            # 2. Step 1: Research & Outline Generation
            logger.info("--> STEP 1: Research & Slide Outline Generation")
            outline = self.researcher.generate_outline(
                topic=input_request.topic,
                num_slides=input_request.num_slides,
                style=input_request.style
            )

            # 3. Step 2: Per-Slide Content Generation
            logger.info("--> STEP 2: Detailed Per-Slide Content Generation")
            content = self.content_gen.generate_presentation_content(
                outline=outline,
                style=input_request.style
            )

            # 4. Step 3: Template Population & Visual Builder
            logger.info("--> STEP 3: Template Population & Layout Rendering")
            builder = PresentationBuilder(template_path=input_request.template)
            built_pptx_path = builder.build_presentation(
                content=content,
                output_pptx_path=pptx_output_path
            )

            # 5. Step 4: Export to PPTX and Convert to PDF
            logger.info("--> STEP 4: Exporting PPTX & Converting to PDF")
            pdf_success, final_pdf_path, conversion_method = self.pdf_exporter.convert_pptx_to_pdf(
                pptx_path=built_pptx_path,
                pdf_output_path=pdf_output_path
            )

            elapsed_time = round(time.time() - start_time, 2)
            logger.info("==========================================================================")
            logger.info(f"PIPELINE COMPLETED SUCCESSFULLY IN {elapsed_time} SECONDS")
            logger.info(f"PPTX Output: {built_pptx_path}")
            if pdf_success:
                logger.info(f"PDF Output: {final_pdf_path} (Method: {conversion_method})")
            logger.info("==========================================================================")

            return PipelineOutput(
                status="SUCCESS",
                topic=input_request.topic,
                num_slides=len(content.slides),
                pptx_path=built_pptx_path,
                pdf_path=final_pdf_path if pdf_success else None,
                pdf_converted=pdf_success,
                conversion_method=conversion_method,
                execution_time_seconds=elapsed_time
            )

        except Exception as e:
            elapsed_time = round(time.time() - start_time, 2)
            logger.error(f"PIPELINE FAILED AFTER {elapsed_time} SECONDS: {e}", exc_info=True)
            return PipelineOutput(
                status="FAILED",
                topic=input_request.topic,
                num_slides=input_request.num_slides,
                pptx_path="",
                error_message=str(e),
                execution_time_seconds=elapsed_time
            )

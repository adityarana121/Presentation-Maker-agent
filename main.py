"""
CLI Entrypoint for AI Presentation Generator Agent.
Supports JSON file input or direct CLI flag arguments.
"""

import sys
import json
import argparse
import logging
from pathlib import Path

from src.config import DEFAULT_TEMPLATE_PATH
from src.models import InputRequest
from src.orchestrator import PresentationAgent

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s"
)
logger = logging.getLogger("Main")


def parse_args():
    parser = argparse.ArgumentParser(
        description="AI Presentation Generator Agent - Tecmotiv Solutions Assessment"
    )
    parser.add_argument(
        "--input", "-i",
        type=str,
        help="Path to JSON input file containing 'template', 'topic', 'num_slides', 'style'"
    )
    parser.add_argument(
        "--topic", "-t",
        type=str,
        help="Presentation topic (e.g. 'Quantum Computing')"
    )
    parser.add_argument(
        "--num_slides", "-n",
        type=int,
        default=8,
        help="Number of slides to generate (default: 8)"
    )
    parser.add_argument(
        "--template",
        type=str,
        default=DEFAULT_TEMPLATE_PATH,
        help="Path to PowerPoint template file (default: TECMOTIV_SOLUTIONS.pptx)"
    )
    parser.add_argument(
        "--style", "-s",
        type=str,
        default="professional",
        help="Presentation style (default: 'professional')"
    )
    parser.add_argument(
        "--provider", "-p",
        type=str,
        default="auto",
        choices=["auto", "groq", "gemini", "openai"],
        help="LLM Provider choice (default: 'auto' preferring Groq)"
    )
    return parser.parse_args()


def main():
    args = parse_args()

    # 1. Parse Input Parameters
    if args.input:
        input_file = Path(args.input)
        if not input_file.exists():
            logger.error(f"Input JSON file not found: {args.input}")
            sys.exit(1)
        with open(input_file, "r", encoding="utf-8") as f:
            data = json.load(f)
        request = InputRequest.model_validate(data)
    elif args.topic:
        request = InputRequest(
            template=args.template,
            topic=args.topic,
            num_slides=args.num_slides,
            style=args.style
        )
    else:
        # Default test run if no arguments provided
        logger.info("No input file or topic specified. Defaulting to sample topic: 'Quantum Computing' (8 slides)")
        request = InputRequest(
            template=DEFAULT_TEMPLATE_PATH,
            topic="Quantum Computing",
            num_slides=8,
            style="professional"
        )

    # 2. Run Presentation Pipeline
    agent = PresentationAgent(llm_provider=args.provider)
    result = agent.run(request)

    # 3. Print Final Summary
    if result.status == "SUCCESS":
        print("\n" + "=" * 60)
        print("[SUCCESS] AI PRESENTATION GENERATION COMPLETED")
        print("=" * 60)
        print(f"Topic:            {result.topic}")
        print(f"Slide Count:      {result.num_slides}")
        print(f"PPTX Output File: {result.pptx_path}")
        print(f"PDF Output File:  {result.pdf_path or 'N/A'}")
        print(f"PDF Method:       {result.conversion_method}")
        print(f"Total Time:       {result.execution_time_seconds} seconds")
        print("=" * 60 + "\n")
        sys.exit(0)
    else:
        print("\n" + "=" * 60)
        print("[FAILED] PRESENTATION GENERATION FAILED")
        print("=" * 60)
        print(f"Error: {result.error_message}")
        print("=" * 60 + "\n")
        sys.exit(1)


if __name__ == "__main__":
    main()

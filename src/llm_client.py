"""
LLM Client Wrapper supporting Groq (primary), Gemini, OpenAI, and Anthropic.
Handles structured JSON output parsing, schema validation, rate limits, and retries.
"""

import os
import re
import json
import time
import logging
from typing import Type, TypeVar, Optional
from pydantic import BaseModel

from src.config import GROQ_API_KEY, GROQ_MODEL, GEMINI_API_KEY, OPENAI_API_KEY, ANTHROPIC_API_KEY

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(name)s: %(message)s")
logger = logging.getLogger("LLMClient")

T = TypeVar("T", bound=BaseModel)


class LLMClient:
    """Unified LLM interface with fallback and structured output parsing."""

    def __init__(self, provider: str = "auto"):
        self.provider = provider
        self.client = None
        self._init_client()

    def _init_client(self):
        """Initialize the client based on available API keys."""
        if (self.provider == "auto" or self.provider == "groq") and GROQ_API_KEY:
            try:
                from groq import Groq
                self.client = Groq(api_key=GROQ_API_KEY)
                self.provider = "groq"
                logger.info(f"Initialized Groq LLM client with model '{GROQ_MODEL}'")
                return
            except Exception as e:
                logger.warning(f"Failed to initialize Groq client: {e}")

        if (self.provider == "auto" or self.provider == "gemini") and GEMINI_API_KEY:
            try:
                from google import genai
                self.client = genai.Client(api_key=GEMINI_API_KEY)
                self.provider = "gemini"
                logger.info("Initialized Google Gemini client")
                return
            except Exception as e:
                logger.warning(f"Failed to initialize Gemini client: {e}")

        if (self.provider == "auto" or self.provider == "openai") and OPENAI_API_KEY:
            try:
                from openai import OpenAI
                self.client = OpenAI(api_key=OPENAI_API_KEY)
                self.provider = "openai"
                logger.info("Initialized OpenAI client")
                return
            except Exception as e:
                logger.warning(f"Failed to initialize OpenAI client: {e}")

        if not self.client:
            raise ValueError(
                "No valid LLM API key configured! Please set GROQ_API_KEY, GEMINI_API_KEY, or OPENAI_API_KEY in .env"
            )

    def generate_json(
        self,
        prompt: str,
        system_prompt: str,
        response_model: Type[T],
        max_retries: int = 3
    ) -> T:
        """
        Generate structured output from the LLM matching response_model schema.
        Includes automatic retry, JSON markdown strip, and schema validation.
        """
        schema_json = json.dumps(response_model.model_json_schema(), indent=2)
        full_system_prompt = (
            f"{system_prompt}\n\n"
            f"CRITICAL REQUIREMENT: You MUST respond ONLY with valid JSON matching this JSON schema:\n"
            f"```json\n{schema_json}\n```\n"
            f"Do not include any conversational text, commentary, or markdown wrapper outside the JSON object."
        )

        for attempt in range(1, max_retries + 1):
            try:
                logger.info(f"Generating structured JSON (Attempt {attempt}/{max_retries}) using provider '{self.provider}'...")
                raw_text = self._call_llm_raw(prompt, full_system_prompt)
                
                # Clean up JSON text (strip ```json ... ``` if present)
                clean_json_str = self._clean_json_markdown(raw_text)
                
                # Parse JSON string
                parsed_dict = json.loads(clean_json_str)
                
                # Validate against Pydantic schema
                validated_obj = response_model.model_validate(parsed_dict)
                logger.info(f"Successfully generated and validated {response_model.__name__} object.")
                return validated_obj

            except Exception as e:
                logger.warning(f"Attempt {attempt}/{max_retries} failed: {e}")
                if attempt == max_retries:
                    logger.error(f"Failed to generate structured JSON after {max_retries} attempts.")
                    raise RuntimeError(f"LLM JSON Generation failed after {max_retries} attempts: {e}")
                time.sleep(1.5 * attempt)  # Exponential backoff

    def _call_llm_raw(self, prompt: str, system_prompt: str) -> str:
        """Direct raw text call to selected provider."""
        if self.provider == "groq":
            response = self.client.chat.completions.create(
                model=GROQ_MODEL,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": prompt}
                ],
                response_format={"type": "json_object"},
                temperature=0.3,
                max_tokens=4096
            )
            return response.choices[0].message.content

        elif self.provider == "gemini":
            response = self.client.models.generate_content(
                model="gemini-2.5-flash",
                contents=prompt,
                config={
                    "system_instruction": system_prompt,
                    "temperature": 0.3,
                    "response_mime_type": "application/json"
                }
            )
            return response.text

        elif self.provider == "openai":
            response = self.client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": prompt}
                ],
                response_format={"type": "json_object"},
                temperature=0.3
            )
            return response.choices[0].message.content

        else:
            raise ValueError(f"Unsupported LLM provider: {self.provider}")

    @staticmethod
    def _clean_json_markdown(text: str) -> str:
        """Strip markdown codeblock backticks and whitespace from JSON response."""
        text = text.strip()
        # Find ```json ... ``` or ``` ... ```
        match = re.search(r"```(?:json)?\s*(.*?)\s*```", text, re.DOTALL)
        if match:
            return match.group(1).strip()
        return text

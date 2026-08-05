# backend/ai_service/llm_router.py
import time
import random
import logging

import ollama
from openai import OpenAI, RateLimitError as OpenAIRateLimitError, APIStatusError as OpenAIAPIStatusError
from anthropic import Anthropic, RateLimitError as AnthropicRateLimitError, APIStatusError as AnthropicAPIStatusError
from django.conf import settings

logger = logging.getLogger(__name__)

# Retry configuration
MAX_RETRIES = 5
BASE_DELAY_SECONDS = 1.0   # first retry waits ~1s, then ~2s, ~4s, ~8s, ~16s (+ jitter)
MAX_DELAY_SECONDS = 30.0


class AIServiceError(Exception):
    """Raised when the AI provider fails for a non-recoverable reason."""
    pass


class AIRateLimitExhausted(AIServiceError):
    """
    Raised specifically when we gave up retrying due to sustained rate limiting.
    Kept distinct from AIServiceError so views.py can return a "try again shortly"
    message instead of a generic failure — this is a capacity/timing issue, not a bug.
    """
    pass


def _is_retryable_status(status_code: int) -> bool:
    """
    429 = rate limited. 500/502/503/504 = transient server-side issues.
    Everything else (400, 401, 403, 404, etc.) is a real problem retrying won't fix.
    """
    return status_code == 429 or status_code >= 500


def _sleep_with_backoff(attempt, retry_after=None):
    """
    Exponential backoff with jitter. If the provider tells us how long to wait
    (Retry-After header, surfaced via retry_after), respect that instead of guessing.
    """
    if retry_after is not None:
        delay = min(retry_after, MAX_DELAY_SECONDS)
    else:
        delay = min(BASE_DELAY_SECONDS * (2 ** attempt), MAX_DELAY_SECONDS)
        delay += random.uniform(0, delay * 0.25)  # jitter, avoids thundering-herd retries

    logger.warning(f"AI request rate-limited/transient error — retrying in {delay:.1f}s (attempt {attempt + 1}/{MAX_RETRIES})")
    time.sleep(delay)


def _extract_retry_after(error):
    """
    Both SDKs expose the original HTTP response on rate-limit errors, which may
    include a Retry-After header. Falls back to None (use our own backoff schedule)
    if it's not present.
    """
    try:
        response = getattr(error, 'response', None)
        if response is not None:
            header_val = response.headers.get('retry-after')
            if header_val is not None:
                return float(header_val)
    except (AttributeError, ValueError, TypeError):
        pass
    return None


class AIClient:
    """
    Unified AI Router.
    Routes prompts to Ollama (Dev), OpenAI, or Anthropic (Claude) based on settings.
    Automatically retries rate-limit (429) and transient server errors with
    exponential backoff; other errors (bad key, invalid request, etc.) fail immediately.
    """

    def __init__(self):
        self.provider = settings.AI_PROVIDER

        if self.provider == 'openai':
            if not settings.OPENAI_API_KEY:
                raise ValueError("OpenAI API Key is missing in .env")
            self.client = OpenAI(api_key=settings.OPENAI_API_KEY)

        elif self.provider == 'anthropic':
            if not settings.ANTHROPIC_API_KEY:
                raise ValueError("Anthropic API Key is missing in .env")
            self.client = Anthropic(api_key=settings.ANTHROPIC_API_KEY)

        elif self.provider == 'ollama':
            self.client = ollama.Client(host=settings.OLLAMA_BASE_URL)

        else:
            raise ValueError(f"Unknown AI Provider: {self.provider}")

    def generate_completion(self, system_prompt: str, user_prompt: str) -> str:
        """
        Sends a prompt to the AI and returns the text response.
        Retries automatically on rate limits / transient server errors.
        """
        last_error = None

        for attempt in range(MAX_RETRIES):
            try:
                return self._call_provider(system_prompt, user_prompt)

            except (AnthropicRateLimitError, OpenAIRateLimitError) as e:
                last_error = e
                retry_after = _extract_retry_after(e)
                if attempt < MAX_RETRIES - 1:
                    _sleep_with_backoff(attempt, retry_after)
                    continue
                break

            except (AnthropicAPIStatusError, OpenAIAPIStatusError) as e:
                status_code = getattr(e, 'status_code', 500)
                last_error = e
                if _is_retryable_status(status_code) and attempt < MAX_RETRIES - 1:
                    _sleep_with_backoff(attempt)
                    continue
                # Non-retryable (400, 401, 403, 404...) — fail immediately, no point retrying
                raise AIServiceError(f"AI Service Error ({self.provider}): {str(e)}")

            except Exception as e:
                # Ollama and unexpected errors don't have structured status codes to
                # check — these aren't retried, since we can't tell transient from permanent.
                raise AIServiceError(f"AI Service Error ({self.provider}): {str(e)}")

        raise AIRateLimitExhausted(
            f"AI Service Error ({self.provider}): gave up after {MAX_RETRIES} attempts "
            f"due to rate limiting. Last error: {str(last_error)}"
        )

    def _call_provider(self, system_prompt: str, user_prompt: str) -> str:
        if self.provider == 'ollama':
            messages = [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ]
            response = self.client.chat(
                model=settings.OLLAMA_MODEL,
                messages=messages,
                options={"temperature": 0.2}
            )
            return response['message']['content']

        elif self.provider == 'openai':
            messages = [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ]
            response = self.client.chat.completions.create(
                model=settings.OPENAI_MODEL,
                messages=messages,
                temperature=0.2
            )
            return response.choices[0].message.content

        elif self.provider == 'anthropic':
            # Anthropic's Messages API takes system prompt as a separate top-level
            # param, not as a message with role "system" (unlike OpenAI/Ollama).
            response = self.client.messages.create(
                model=settings.ANTHROPIC_MODEL,
                max_tokens=4096,
                temperature=0.2,
                system=system_prompt,
                messages=[
                    {"role": "user", "content": user_prompt}
                ]
            )
            return "".join(
                block.text for block in response.content if block.type == "text"
            )


# Instantiate a global client for easy importing
ai_client = AIClient()
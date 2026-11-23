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

# Bug fix (Nov 2026): was 16384. Sonnet 5's adaptive thinking (see the
# temperature comment in _call_provider below — it can't be dialed down or
# disabled per-request) occasionally burned through the ENTIRE max_tokens
# budget reasoning about a dense OM/T12/rent-roll extraction, leaving zero
# tokens for the actual JSON answer and surfacing as "The AI returned an
# empty response" in production. Doubling the budget gives much more room
# for both the reasoning and the (often large, 47+ field) final JSON to fit
# in the same response.
ANTHROPIC_MAX_TOKENS = 32768


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


class AIEmptyResponseError(AIServiceError):
    """
    Raised when the Anthropic response comes back with no usable text content
    -- observed in production with Sonnet 5's adaptive thinking occasionally
    consuming the entire max_tokens budget on reasoning (or, less commonly,
    the request being refused outright), leaving nothing for the actual
    answer. Kept distinct from a plain AIServiceError for two reasons: (1)
    generate_completion's retry loop below treats it like a transient
    condition worth retrying — how much the model chooses to "think" before
    answering isn't fully deterministic, so a second attempt on the exact
    same prompt often succeeds where the first didn't; (2) if every retry
    still comes back empty, views.py can tell this apart from a genuine
    "the AI gave us garbage" failure and respond the same way it already
    does for AIRateLimitExhausted — a retryable 503, NOT a 400 that wrongly
    blames the user's uploaded documents, and NOT recorded as a failed
    analysis in History.
    """
    pass


def _is_retryable_status(status_code: int) -> bool:
    """
    429 = rate limited. 500/502/503/504 = transient server-side issues.
    Everything else (400, 401, 403, 404, etc.) is a real problem retrying won't fix.
    """
    return status_code == 429 or status_code >= 500


def _sleep_with_backoff(attempt, retry_after=None, reason="rate-limited/transient error"):
    """
    Exponential backoff with jitter. If the provider tells us how long to wait
    (Retry-After header, surfaced via retry_after), respect that instead of guessing.

    `reason` is purely cosmetic — it's folded into the log line below so a
    retry triggered by something other than rate limiting (e.g.
    AIEmptyResponseError) doesn't get mislabeled as one in the logs. Defaults
    to the original wording so the existing rate-limit/APIStatusError call
    sites below don't need to change.
    """
    if retry_after is not None:
        delay = min(retry_after, MAX_DELAY_SECONDS)
    else:
        delay = min(BASE_DELAY_SECONDS * (2 ** attempt), MAX_DELAY_SECONDS)
        delay += random.uniform(0, delay * 0.25)  # jitter, avoids thundering-herd retries

    logger.warning(f"AI request {reason} — retrying in {delay:.1f}s (attempt {attempt + 1}/{MAX_RETRIES})")
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

            except AIEmptyResponseError as e:
                # How much the model chooses to "think" before answering isn't
                # fully deterministic for the same prompt — a retry genuinely
                # has a real chance of succeeding where this attempt didn't,
                # unlike a plain empty string we'd have no reason to retry.
                last_error = e
                if attempt < MAX_RETRIES - 1:
                    _sleep_with_backoff(
                        attempt,
                        reason="returned no text content (adaptive thinking likely consumed the full token budget)"
                    )
                    continue
                raise

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
            # Sonnet 5 deprecated manual sampling parameters — setting temperature
            # to any non-default value now returns a 400 error, since the model
            # uses adaptive thinking by default instead of manual sampling control.
            # Omit it entirely rather than pass any value.
            #
            # Bug fix (Nov 2026): a plain (non-streaming) .create() call with
            # ANTHROPIC_MAX_TOKENS this large (raised from 16384 to fix the
            # empty-response issue — see that constant's own comment) trips
            # the Anthropic Python SDK's own client-side guard: "Streaming is
            # required for operations that may take longer than 10 minutes."
            # A long-held blocking HTTP request is exactly what times out on
            # proxies/load balancers along the way, which is why the SDK
            # refuses to even attempt it synchronously past that threshold.
            # .stream()'s get_final_message() blocks until the full response
            # arrives and hands back the same Message shape .create() would
            # have, so nothing below (text extraction, stop_reason) changes —
            # this only changes HOW the response is fetched, not what it is.
            with self.client.messages.stream(
                model=settings.ANTHROPIC_MODEL,
                max_tokens=ANTHROPIC_MAX_TOKENS,
                system=system_prompt,
                messages=[
                    {"role": "user", "content": user_prompt}
                ]
            ) as stream:
                response = stream.get_final_message()
            text = "".join(
                block.text for block in response.content if block.type == "text"
            )
            if not text.strip():
                # Nothing usable came back — most likely adaptive thinking
                # consumed the entire max_tokens budget before it could write
                # an answer, less commonly the request was refused outright.
                # Raise rather than silently return "" so generate_completion's
                # retry loop gets a chance to try again (see AIEmptyResponseError's
                # own docstring for why a retry is actually worth it here).
                raise AIEmptyResponseError(
                    f"AI Service Error ({self.provider}): model returned no text "
                    f"content (stop_reason={getattr(response, 'stop_reason', 'unknown')})."
                )
            return text


# Instantiate a global client for easy importing
ai_client = AIClient()
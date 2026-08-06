# backend/ai_service/analysis_cache.py
"""
Short-lived server-side cache so "Analyze" and "Download Excel" don't each
independently re-parse files and re-call the AI for what the user experiences
as one action. See conversation history for the full problem description.

CAVEAT: uses Django's default cache framework. Works out of the box with the
default in-memory cache backend for single-process dev. In production with
multiple worker processes, in-memory cache is NOT shared across processes —
you'll need a shared backend (Redis) for this to work reliably at scale. This
is the same Redis already recommended for Celery in the scaling discussion.
"""
import uuid
from django.core.cache import cache

ANALYSIS_CACHE_TTL_SECONDS = 30 * 60  # 30 minutes — enough time to review the
                                        # preview modal and then download


def store_analysis(metrics: dict, rent_roll_df) -> str:
    """
    Caches an analysis result and returns a short-lived ID the frontend can
    pass to the download endpoint to reuse this exact result.
    """
    analysis_id = str(uuid.uuid4())
    cache.set(
        f"analysis:{analysis_id}",
        {"metrics": metrics, "rent_roll_df": rent_roll_df},
        timeout=ANALYSIS_CACHE_TTL_SECONDS,
    )
    return analysis_id


def get_cached_analysis(analysis_id: str):
    """
    Returns (metrics, rent_roll_df) if the ID is valid and not expired,
    otherwise None. Callers should treat None as "fall back to reprocessing",
    not as an error.
    """
    if not analysis_id:
        return None

    data = cache.get(f"analysis:{analysis_id}")
    if not data:
        return None

    return data["metrics"], data["rent_roll_df"]
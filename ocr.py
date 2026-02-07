"""
macOS Vision OCR for OneNote embedded images.

Uses pyobjc (Vision + Quartz frameworks) to perform text recognition.
Falls back gracefully on non-macOS platforms.

OCR results are cached on disk keyed by SHA-256 of image bytes to avoid
redundant recognition across backup rescans.
"""

import hashlib
import json
import logging
import sys
from pathlib import Path

log = logging.getLogger("onenote-mcp")

OCR_CACHE_DIR = Path.home() / ".cache" / "onenote-mcp" / "ocr"

_VISION_AVAILABLE = False

if sys.platform == "darwin":
    try:
        import Quartz
        import Vision

        _VISION_AVAILABLE = True
    except ImportError:
        log.warning(
            "pyobjc-framework-Vision/Quartz not installed — OCR disabled. "
            "Install with: pip install pyobjc-framework-Vision pyobjc-framework-Quartz"
        )


def _cache_key(image_bytes: bytes) -> str:
    return hashlib.sha256(image_bytes).hexdigest()


def _load_cached(key: str) -> str | None:
    cache_file = OCR_CACHE_DIR / f"{key}.json"
    if cache_file.exists():
        try:
            data = json.loads(cache_file.read_text(encoding="utf-8"))
            return data.get("text", "")
        except Exception:
            return None
    return None


def _save_cache(key: str, text: str) -> None:
    try:
        OCR_CACHE_DIR.mkdir(parents=True, exist_ok=True)
        cache_file = OCR_CACHE_DIR / f"{key}.json"
        cache_file.write_text(
            json.dumps({"text": text}, ensure_ascii=False),
            encoding="utf-8",
        )
    except Exception as e:
        log.debug("OCR cache write failed: %s", e)


def ocr_image(image_bytes: bytes) -> str:
    """Run macOS Vision OCR on image bytes, return recognized text.

    Returns empty string if OCR is unavailable or no text is found.
    Results are cached on disk keyed by SHA-256 of the image bytes.
    """
    key = _cache_key(image_bytes)
    cached = _load_cached(key)
    if cached is not None:
        log.debug("OCR cache hit: %s", key[:12])
        return cached

    if not _VISION_AVAILABLE:
        return ""

    try:
        data = Quartz.CFDataCreate(None, image_bytes, len(image_bytes))
        image_source = Quartz.CGImageSourceCreateWithData(data, None)
        if image_source is None:
            log.debug("OCR: could not create image source")
            return ""

        cg_image = Quartz.CGImageSourceCreateImageAtIndex(image_source, 0, None)
        if cg_image is None:
            log.debug("OCR: could not create CGImage from source")
            return ""

        request = Vision.VNRecognizeTextRequest.alloc().init()
        request.setRecognitionLevel_(Vision.VNRequestTextRecognitionLevelAccurate)
        request.setRecognitionLanguages_(["en", "ru"])
        request.setUsesLanguageCorrection_(True)

        handler = Vision.VNImageRequestHandler.alloc().initWithCGImage_options_(
            cg_image, None
        )
        success = handler.performRequests_error_([request], None)
        if not success[0]:
            log.debug("OCR: VNImageRequestHandler failed: %s", success[1])
            return ""

        results = request.results()
        if not results:
            return ""

        lines = []
        for observation in results:
            candidate = observation.topCandidates_(1)
            if candidate:
                lines.append(candidate[0].string())

        text = "\n".join(lines)
        _save_cache(key, text)
        return text

    except Exception as e:
        log.warning("OCR failed: %s", e)
        return ""

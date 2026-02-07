"""
macOS Vision OCR for OneNote embedded images.

Uses pyobjc (Vision + Quartz frameworks) to perform text recognition.
Falls back gracefully on non-macOS platforms.
"""

import logging
import sys

log = logging.getLogger("onenote-mcp")

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


def ocr_image(image_bytes: bytes) -> str:
    """Run macOS Vision OCR on image bytes, return recognized text.

    Returns empty string if OCR is unavailable or no text is found.
    """
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

        return "\n".join(lines)

    except Exception as e:
        log.warning("OCR failed: %s", e)
        return ""

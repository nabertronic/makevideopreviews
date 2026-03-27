class MakeVideoPreviewsError(Exception):
    """Base application error."""


class ValidationError(MakeVideoPreviewsError):
    """Raised when user input is invalid."""


class DependencyError(MakeVideoPreviewsError):
    """Raised when a required dependency is unavailable."""


class ProbeError(MakeVideoPreviewsError):
    """Raised when ffprobe metadata cannot be read."""


class FrameExtractionError(MakeVideoPreviewsError):
    """Raised when ffmpeg frame extraction fails."""

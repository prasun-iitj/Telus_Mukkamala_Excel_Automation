"""Custom exceptions for the Excel-to-PPT automation pipeline."""


class PipelineError(Exception):
    """Base exception for pipeline failures."""

    def __init__(self, step: str, message: str):
        self.step = step
        super().__init__(f"[{step}] {message}")


class SheetNotFoundError(PipelineError):
    """Raised when a required Excel sheet is missing."""

    def __init__(self, sheet_name: str):
        self.sheet_name = sheet_name
        super().__init__("validation", f"Required sheet '{sheet_name}' not found")


class ShapeNotFoundError(PipelineError):
    """Raised when a required shape is not found on a slide."""

    def __init__(self, slide_index: int, shape_name: str):
        self.slide_index = slide_index
        self.shape_name = shape_name
        super().__init__("ppt_update", f"Shape '{shape_name}' not found on slide {slide_index}")


class SlideMapError(PipelineError):
    """Raised when the slide map YAML configuration is invalid."""

    def __init__(self, message: str):
        super().__init__("config", message)

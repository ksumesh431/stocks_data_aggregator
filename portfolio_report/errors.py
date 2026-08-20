"""Application-specific errors that are safe to show without a traceback."""


class WorkbookOutputError(RuntimeError):
    """Raised when the destination workbook cannot be created or replaced."""

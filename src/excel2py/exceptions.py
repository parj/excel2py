"""Custom exceptions for excel2py."""


class Excel2PyError(Exception):
    """Base exception for excel2py."""


class ExcelParseError(Excel2PyError):
    """Raised when an Excel file cannot be parsed."""


class UnsupportedFormatError(Excel2PyError):
    """Raised for unsupported Excel file formats."""


class ProviderError(Excel2PyError):
    """Base exception for LLM provider errors."""


class ProviderAuthError(ProviderError):
    """Raised when LLM provider authentication fails."""


class ProviderRateLimitError(ProviderError):
    """Raised when LLM provider rate limit is hit."""


class CodeGenerationError(Excel2PyError):
    """Raised when generated code is invalid."""

"""User-facing errors shared by document, cloud, agent, and HTTP layers."""


class DemoError(Exception):
    """Carry a safe user-visible message and the intended HTTP response status."""

    def __init__(self, message: str, status: int = 400) -> None:
        """Initialize the exception message and HTTP status without attaching secrets."""
        super().__init__(message)
        self.message = message
        self.status = status

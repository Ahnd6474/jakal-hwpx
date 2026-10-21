class WorkspaceError(ValueError):
    """A transport-independent, safe-to-display document operation failure."""

    def __init__(self, code: str, message: str, status: int = 409):
        super().__init__(message)
        self.code = code
        self.status = status

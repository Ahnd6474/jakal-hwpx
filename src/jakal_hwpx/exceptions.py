class HwpxError(Exception):
    pass


class InvalidHwpxFileError(HwpxError):
    pass


class HwpxValidationError(HwpxError):
    def __init__(self, errors: list[str]):
        super().__init__("\n".join(errors))
        self.errors = errors

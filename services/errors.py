class ServiceError(Exception):
    """Error that is safe to show to the user."""

    def __init__(self, message, title="Error"):
        super().__init__(message)
        self.title = title
        self.message = message


class DuplicateCircuitError(ServiceError):
    def __init__(self, circuits):
        duplicate_list = ", ".join(str(c) for c in circuits)
        super().__init__(
            f"Circuit No duplicate di Lokasi {duplicate_list}! Tolong cek BOM kembali!",
            title="Warning",
        )
        self.circuits = circuits


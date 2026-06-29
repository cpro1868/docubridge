from dataclasses import dataclass
from enum import StrEnum


class Severity(StrEnum):
    INFO = "info"
    WARNING = "warning"
    ERROR = "error"
    FATAL = "fatal"


@dataclass(slots=True)
class Diagnostic:
    severity: Severity
    code: str
    message: str
    location: str | None = None
    hint: str | None = None

    def to_dict(self) -> dict[str, str | None]:
        return {
            "severity": self.severity.value,
            "code": self.code,
            "message": self.message,
            "location": self.location,
            "hint": self.hint,
        }

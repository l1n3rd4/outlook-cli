"""Contact model."""

from __future__ import annotations

from dataclasses import dataclass

from .common import EmailAddress


@dataclass
class Contact:
    id: str
    display_name: str
    given_name: str
    surname: str
    email_addresses: list[EmailAddress]
    company: str
    job_title: str

    @classmethod
    def from_api(cls, data: dict) -> Contact:
        emails = [
            EmailAddress(name=e.get("Name", ""), address=e.get("Address", ""))
            for e in data.get("EmailAddresses", [])
        ]
        return cls(
            id=data["Id"],
            display_name=data.get("DisplayName", ""),
            given_name=data.get("GivenName", ""),
            surname=data.get("Surname", ""),
            email_addresses=emails,
            company=data.get("CompanyName", ""),
            job_title=data.get("JobTitle", ""),
        )

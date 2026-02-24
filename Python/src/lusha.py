"""
lusha.py — Lusha API integratie

Verantwoordelijkheden:
  - Contacten zoeken via /prospecting/contact/search/
  - Contacten enrichen via /prospecting/contact/enrich
  - Industrieën ophalen via /prospecting/filters/companies/industries_labels
  - Pagination afhandelen
"""

from __future__ import annotations

import time
from typing import Optional

import requests

BASE_URL = "https://api.lusha.com/prospecting"


class LushaClient:
    def __init__(self, api_key: str):
        if not api_key:
            raise ValueError("Lusha API key ontbreekt. Zet LUSHA_API_KEY in .env")
        self.api_key = api_key
        self.session = requests.Session()
        self.session.headers.update({"api_key": api_key, "Content-Type": "application/json"})
        self._last_request_id: Optional[str] = None

    # ── Industries ────────────────────────────────────────────────────────

    def get_industries(self) -> list[dict]:
        """
        Haal alle beschikbare industrieën op van Lusha.

        Returns:
            Lijst van dicts met keys: main_industry, main_industry_id, sub_industries
        """
        resp = self.session.get(
            f"{BASE_URL}/filters/companies/industries_labels",
            timeout=30,
        )
        resp.raise_for_status()
        return resp.json()

    # ── Search ────────────────────────────────────────────────────────────

    def search_contacts(
        self,
        page: int = 1,
        page_size: int = 10,
        countries: list[str] = None,
        company_sizes: list[dict] = None,
        industry_ids: list[int] = None,
        job_titles: list[str] = None,
    ) -> dict:
        payload = {
            "pages": {"page": page, "size": page_size},
            "filters": {
                "companies": {
                    "include": {
                        "locations": [{"country": c} for c in (countries or ["Netherlands"])],
                        "sizes": company_sizes or [{"min": 51, "max": 1000}],
                        **({"mainIndustriesIds": industry_ids} if industry_ids else {}),
                    }
                },
                "contacts": {
                    "include": {
                        "jobTitles": job_titles or [
                            "CEO", "Chief Executive Officer",
                            "CFO", "Chief Financial Officer",
                            "COO", "Chief Operating Officer",
                            "CTO", "Chief Technology Officer",
                            "CMO", "Chief Marketing Officer",
                        ]
                    }
                },
            },
        }

        resp = self.session.post(f"{BASE_URL}/contact/search/", json=payload, timeout=30)
        resp.raise_for_status()
        data = resp.json()

        if "error" in data:
            raise RuntimeError(f"Lusha search error: {data['error']}")

        request_id = data.get("requestId", "")
        contacts   = data.get("data", [])
        total      = data.get("total", len(contacts))
        self._last_request_id = request_id

        print(f"[LUSHA] Search pagina {page}: {len(contacts)} contacten gevonden "
              f"(totaal: {total}, requestId: {request_id})")

        return {
            "contacts":   contacts,
            "request_id": request_id,
            "total":      total,
            "page":       page,
        }

    def search_multiple_pages(
        self,
        num_pages: int = 1,
        page_size: int = 10,
        start_page: int = 1,
        **kwargs,
    ) -> tuple[list[dict], str]:
        all_contacts = []
        last_request_id = ""

        for page in range(start_page, start_page + num_pages):
            result = self.search_contacts(page=page, page_size=page_size, **kwargs)
            all_contacts.extend(result["contacts"])
            last_request_id = result["request_id"]

            if not result["contacts"]:
                print(f"[LUSHA] Geen resultaten meer op pagina {page}, stoppen.")
                break

            if page < start_page + num_pages - 1:
                time.sleep(0.5)

        return all_contacts, last_request_id

    # ── Enrich ────────────────────────────────────────────────────────────

    def enrich_contacts(self, request_id: str, contact_ids: list[str]) -> list[dict]:
        if not request_id:
            raise ValueError("request_id is verplicht voor enrichment")
        if not contact_ids:
            return []

        payload = {"requestId": request_id, "contactIds": contact_ids}
        resp = self.session.post(f"{BASE_URL}/contact/enrich", json=payload, timeout=30)
        resp.raise_for_status()
        data = resp.json()

        if "error" in data:
            raise RuntimeError(f"Lusha enrich error: {data['error']}")

        raw_contacts = data.get("contacts", [])
        return [self._parse_enriched(c) for c in raw_contacts]

    @staticmethod
    def _parse_enriched(contact: dict) -> dict:
        contact_data = contact.get("data", {})
        contact_id   = str(contact.get("id") or contact.get("contactId") or "")

        emails  = contact_data.get("emailAddresses", [])
        phones  = contact_data.get("phoneNumbers", [])
        social  = contact_data.get("socialLinks", {})

        return {
            "contact_id": contact_id,
            "email":      emails[0]["email"] if emails else "",
            "phone":      phones[0]["number"] if phones else "",
            "linkedin":   social.get("linkedin", ""),
            "all_emails": [e["email"] for e in emails],
            "all_phones": [p["number"] for p in phones],
        }


# ── ICP presets ───────────────────────────────────────────────────────────

ICP_PRESETS = {
    "nl_midsized_csuite": {
        "countries":      ["Netherlands"],
        "company_sizes":  [{"min": 51, "max": 1000}],
        "industry_ids":   [],   # leeg = alle industrieën (of stel in via menu)
        "job_titles": [
            "CEO", "Chief Executive Officer",
            "CFO", "Chief Financial Officer",
            "COO", "Chief Operating Officer",
            "CTO", "Chief Technology Officer",
            "CMO", "Chief Marketing Officer",
        ],
    },
    "nl_large_csuite": {
        "countries":      ["Netherlands"],
        "company_sizes":  [{"min": 1001, "max": 10000}],
        "industry_ids":   [],
        "job_titles": [
            "CEO", "CFO", "COO", "CTO", "CMO",
            "Director", "Managing Director",
        ],
    },
}
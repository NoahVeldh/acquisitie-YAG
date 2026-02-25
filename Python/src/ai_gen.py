"""
ai_gen.py — AI bericht generatie via OpenAI

Verantwoordelijkheden:
  - Gepersonaliseerde acquisitie e-mails genereren per lead
  - Vaste blokken (opening, pitch, signature) samenvoegen met AI-connectiezinnen
  - Web search via OpenAI om bedrijfsinformatie op te zoeken
  - Foutafhandeling bij ontbrekende bedrijfsinfo
  - Dry-run modus (preview zonder API calls)

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
WIJZIGING T.O.V. VORIGE VERSIE:
  Prompt herschreven voor persoonlijker resultaat:
    - Ontvanger wordt direct aangesproken (jij/jouw/jullie), nooit meer "hun/het bedrijf"
    - Connectiezinnen starten vanuit de uitdaging/ambitie van de ontvanger, niet
      vanuit Rick's persoonlijke interesse
    - "Wij/we" voor YAG-capabilities, "ik" alleen voor persoonlijke noot van Rick
    - C-suite toon: direct en bondig, geen lof of vleierij

EERDERE WIJZIGINGEN:
  URL-stripping toegevoegd als post-processing stap
    - _strip_urls() verwijdert alle URLs en losse domeinnamen uit de gegenereerde tekst
    - Pakt: https://..., http://..., www.bedrijf.nl, bedrijf.com/pad etc.
    - Wordt aangeroepen direct na _generate_connection_sentences()
    - Dubbele spaties en spaties voor leestekens worden automatisch opgeruimd
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
"""
from __future__ import annotations

import re
from openai import OpenAI


# ── Vaste e-mail blokken ──────────────────────────────────────────────────

DEFAULT_OPENING = (
    "Ik ben {sender_name}, student {studie} aan de {universiteit} "
    "en werkzaam bij Young Advisory Group (YAG), een volledig studenten-gerund adviesbureau. "
    "Omdat wij onze eigen acquisitie doen, ben ik actief op zoek naar bedrijven "
    "waar onze kennis waarde kan toevoegen."
)

DEFAULT_PITCH = """\
Met de Young Advisory Group (YAG) hebben we door de jaren heen veel verschillende projecten \
afgerond in veel verschillende industrieën. Om concrete voorbeelden te geven: recent hebben we \
gewerkt aan een project voor een bedrijf in de biertank industrie waarbij we geholpen hebben \
met de implementatie van een ERP systeem. Ook hebben we advies gegeven aan een grote dierentuin \
hoe zij het beste 'dynamic pricing' kon implementeren. Zelf ben ik nu bezig met een project \
voor de GGD waarin we gepersonaliseerde brieven opstellen voor 35 gemeenteraden om het belang \
van gezond opgroeien te benadrukken.

Graag zou ik willen voorstellen om een gesprek in te plannen, waarin we kennis kunnen maken \
en de mogelijkheden voor een eventuele samenwerking kunnen verkennen.

Ik hoor graag of dit schikt en bij verdere vragen ben ik altijd bereikbaar!\
"""

DEFAULT_SIGNATURE = """\
Met vriendelijke groet,

{sender_name}
Strategy Consultant - Young Advisory Group
–––––––––––––––––––––––––
Vestiging {vestiging}
Videolab
Torenallee 20
5617 BC Eindhoven
{sender_phone}
{sender_email} | LinkedIn | www.youngadvisorygroup.nl"""


# ── URL stripper ──────────────────────────────────────────────────────────

# Patronen die worden verwijderd:
#   1. Volledige URLs: https://..., http://...
#   2. www.iets.tld (met of zonder pad)
#   3. Losse domeinnamen: bedrijf.nl, bedrijf.com/pad (minimaal 2-char tld)
#      — alleen als ze NIET in de vaste signature staan (die wordt na assembly toegevoegd)
_URL_PATTERN = re.compile(
    r"(?:"
    r"https?://[^\s\)\]\,]+"           # http(s)://...
    r"|www\.[^\s\)\]\,]+"              # www.iets.nl
    r"|[a-zA-Z0-9\-]+\.[a-zA-Z]{2,}"  # bedrijf.nl / bedrijf.com
    r"(?:/[^\s\)\]\,]*)?"              # optioneel pad
    r")",
    re.IGNORECASE,
)

# Leestekens die na verwijdering direct naast een spatie komen te staan
_PUNCT_SPACE = re.compile(r"\s+([,\.;:!?])")


def _strip_urls(text: str) -> str:
    """
    Verwijder alle URLs en losse domeinnamen uit tekst.
    Ruimt daarna ook markdown-link skeletten op die overblijven na URL-verwijdering,
    zoals ([.com]()), ([tekst]()) of losse lege haakjes.
    Ruimt daarna dubbele spaties en spaties voor leestekens op.
    """
    cleaned = _URL_PATTERN.sub("", text)
    cleaned = re.sub(r"\(\[.*?\]\(\s*\)\)", "", cleaned)   # ([tekst]())
    cleaned = re.sub(r"\[.*?\]\(\s*\)", "", cleaned)        # [tekst]()
    cleaned = re.sub(r"\(\s*\)", "", cleaned)               # lege () haakjes
    cleaned = _PUNCT_SPACE.sub(r"\1", cleaned)              # spatie voor leesteken weg
    cleaned = re.sub(r" {2,}", " ", cleaned)                # dubbele spaties -> een
    cleaned = "\n".join(line.strip() for line in cleaned.splitlines())
    return cleaned.strip()


# ── AI Client ─────────────────────────────────────────────────────────────

class AIGenerator:
    def __init__(
        self,
        api_key: str,
        model: str = "gpt-4.1-mini",
        sender_name: str = "",
        sender_email: str = "",
        sender_phone: str = "",
        sender_linkedin: str = "www.youngadvisorygroup.nl",
        studie: str = "Technische Bedrijfskunde",
        universiteit: str = "TU Eindhoven",
        opening_template: str = DEFAULT_OPENING,
        pitch: str = DEFAULT_PITCH,
        signature_template: str = DEFAULT_SIGNATURE,
        use_web_search: bool = True,
    ):
        if not api_key:
            raise ValueError("OpenAI API key ontbreekt. Zet OPENAI_API_KEY in .env")
        self.client             = OpenAI(api_key=api_key)
        self.model              = model
        self.sender             = sender_name
        self.email              = sender_email
        self.phone              = sender_phone
        self.linkedin           = sender_linkedin
        self.studie             = studie
        self.universiteit       = universiteit
        self.use_web_search     = use_web_search
        self.opening_template   = opening_template
        self.pitch              = pitch
        self.signature_template = signature_template

    # ── Publieke interface ────────────────────────────────────────────────

    def generate(
        self,
        first_name: str,
        job_title: str,
        company_name: str,
        website: str,
        vestiging: str = "",
    ) -> tuple[str, int]:
        """
        Genereer een volledig e-mailbericht voor één lead.
        Als use_web_search=True zoekt OpenAI zelf op wat het bedrijf doet.
        URLs worden na generatie automatisch uit de connectiezinnen verwijderd.

        Returns:
            (bericht, tokens) — de volledige mail-tekst en het totale tokenverbruik
        """
        if not company_name:
            raise ValueError("company_name is verplicht voor AI generatie")
        if not first_name:
            raise ValueError("first_name is verplicht voor AI generatie")

        connection_text, tokens = self._generate_connection_sentences(
            first_name=first_name,
            job_title=job_title,
            company_name=company_name,
            website=website,
        )

        # Verwijder URLs die het model toch heeft toegevoegd
        connection_text = _strip_urls(connection_text)

        bericht = self._assemble_email(
            first_name=first_name,
            company_name=company_name,
            connection_text=connection_text,
            vestiging=vestiging,
        )
        return bericht, tokens

    def preview(
        self,
        first_name: str,
        company_name: str,
        connection_placeholder: str = "[AI CONNECTIEZINNEN KOMEN HIER]",
        vestiging: str = "",
    ) -> tuple[str, int]:
        """
        Genereer een preview zonder API call (dry-run).

        Returns:
            (bericht, 0) — tokens is altijd 0 bij een preview
        """
        bericht = self._assemble_email(
            first_name=first_name,
            company_name=company_name,
            connection_text=connection_placeholder,
            vestiging=vestiging,
        )
        return bericht, 0

    # ── Interne methodes ──────────────────────────────────────────────────

    def _generate_connection_sentences(
        self,
        first_name: str,
        job_title: str,
        company_name: str,
        website: str,
    ) -> tuple[str, int]:
        """
        Vraag OpenAI om 2-3 gepersonaliseerde connectiezinnen.
        Met web search zoekt het model eerst op wat het bedrijf doet.

        Returns:
            (tekst, tokens) — de connectiezinnen en het totale tokenverbruik
        """
        website_hint = f"De website is {website}." if website else ""
        sender_first = self.sender.split()[0]

        prompt = f"""
Zoek op wat {company_name} doet en welke uitdagingen of strategische ambities relevant zijn \
voor {first_name} als {job_title}. {website_hint}

Schrijf precies 2-3 zinnen in vloeiend Nederlands, direct gericht aan {first_name} persoonlijk. \
{sender_first} schrijft namens Young Advisory Group (YAG), een studenten-adviesbureau.

Gewenste structuur:
1. Eerste zin: {sender_first} vertelt wat hem persoonlijk aansprak aan {company_name} — \
   iets concreets wat hij tegenkwam tijdens zijn studie {self.studie} \
   (denk aan logistiek, operations, strategie, data, procesoptimalisatie). \
   Dit is een persoonlijke noot van {sender_first}, dus gebruik "ik" en wees specifiek.
2. Tweede/derde zin: maak de brug naar wat YAG kan betekenen voor {first_name} \
   — gebruik "wij/we" voor YAG's aanpak. Spreek {first_name} direct aan \
   met "jij/jouw/jullie".

Harde regels:
- Spreek {first_name} altijd direct aan: "jij", "jouw", "jullie" — \
  NOOIT "het bedrijf", "hen", "hun" of "{company_name}" als derde persoon
- Geen vleierij: niet beginnen met "Ik ben gefascineerd door..." of \
  "Ik volg jullie met interesse"
- Geen URLs, geen haakjes met links
- Geen woorden als "innovatief", "onder de indruk", "disruptief"
- Vat NIET samen wat {company_name} doet — {first_name} weet dat zelf
- Toon: direct en bondig, geschikt voor C-suite

Geef ALLEEN de 2-3 zinnen terug, niets anders.
""".strip()

        tools = [{"type": "web_search_preview"}] if self.use_web_search else []

        response = self.client.responses.create(
            model=self.model,
            tools=tools if tools else None,
            input=[{"role": "user", "content": [{"type": "input_text", "text": prompt}]}],
        )

        # Lees tokenverbruik uit de response
        usage = getattr(response, "usage", None)
        if usage:
            tokens = getattr(usage, "input_tokens", 0) + getattr(usage, "output_tokens", 0)
        else:
            tokens = 0

        # Parse output tekst
        output_text = ""
        for item in (response.output or []):
            if getattr(item, "type", "") == "message":
                for c in (item.content or []):
                    if getattr(c, "type", "") == "output_text":
                        output_text = c.text
                        break

        # Fallback voor oudere response structuur
        if not output_text:
            output_text = getattr(response, "output_text", "") or ""

        if not output_text.strip():
            raise RuntimeError("OpenAI gaf een lege response terug")

        return output_text.strip(), tokens

    def _assemble_email(
        self,
        first_name: str,
        company_name: str,
        connection_text: str,
        vestiging: str = "",
    ) -> str:
        """Stel het volledige e-mailbericht samen uit de vaste blokken + AI tekst."""
        opening = self.opening_template.format(
            sender_name=self.sender,
            studie=self.studie,
            universiteit=self.universiteit,
        )
        signature = self.signature_template.format(
            sender_name=self.sender,
            sender_email=self.email,
            sender_phone=self.phone,
            sender_linkedin=self.linkedin,
            vestiging=vestiging or "Eindhoven-Tilburg",
        )
        lines = [
            f"Beste {first_name},",
            "",
            f"{opening} {connection_text} Graag verken ik de mogelijkheden voor een samenwerking.",
            "",
            self.pitch,
            "",
            signature,
        ]
        return "\n".join(lines)

    @staticmethod
    def subject(company_name: str, template: str = "Young Advisory Group x {company}") -> str:
        """Genereer het e-mailonderwerp."""
        return template.format(company=company_name or "jullie")
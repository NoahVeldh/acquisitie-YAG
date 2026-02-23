"""
ai_gen.py — AI bericht generatie via OpenAI

Verantwoordelijkheden:
  - Gepersonaliseerde acquisitie e-mails genereren per lead
  - Vaste blokken (opening, pitch, signature) samenvoegen met AI-connectiezinnen
  - Foutafhandeling bij ontbrekende bedrijfsinfo
  - Dry-run modus (preview zonder API calls)
"""

from __future__ import annotations

from openai import OpenAI


# ── Vaste e-mail blokken ──────────────────────────────────────────────────
# Pas hier de tekst aan voor andere consultants via .env of via parameters

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

Ik hoor graag of dit schikt en bij verdere vragen ben ik altijd bereikbaar!"""

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
{sender_email} | Linkedin www.youngadvisorygroup.nl"""


# ── AI Client ─────────────────────────────────────────────────────────────

class AIGenerator:
    def __init__(
        self,
        api_key: str,
        model: str = "gpt-4.1-mini",
        sender_name: str = "",
        sender_email: str = "",
        sender_phone: str = "",
        studie: str = "Technische Bedrijfskunde",
        universiteit: str = "TU Eindhoven",
        opening_template: str = DEFAULT_OPENING,
        pitch: str = DEFAULT_PITCH,
        signature_template: str = DEFAULT_SIGNATURE,
    ):
        if not api_key:
            raise ValueError("OpenAI API key ontbreekt. Zet OPENAI_API_KEY in .env")

        self.client     = OpenAI(api_key=api_key)
        self.model      = model
        self.sender     = sender_name
        self.email      = sender_email
        self.phone      = sender_phone
        self.studie     = studie
        self.universiteit = universiteit

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
    ) -> str:
        """
        Genereer een volledig e-mailbericht voor één lead.

        Returns:
            Volledig opgemaakt e-mailbericht als string.

        Raises:
            ValueError: Als verplichte velden ontbreken.
            RuntimeError: Als de OpenAI API call mislukt.
        """
        if not company_name:
            raise ValueError("company_name is verplicht voor AI generatie")
        if not first_name:
            raise ValueError("first_name is verplicht voor AI generatie")

        # Stap 1: genereer de connectiezinnen via AI
        connection_text = self._generate_connection_sentences(
            first_name=first_name,
            job_title=job_title,
            company_name=company_name,
            website=website,
        )

        # Stap 2: stel het volledige bericht samen
        return self._assemble_email(
            first_name=first_name,
            company_name=company_name,
            connection_text=connection_text,
            vestiging=vestiging,
        )

    def preview(
        self,
        first_name: str,
        company_name: str,
        connection_placeholder: str = "[AI CONNECTIEZINNEN KOMEN HIER]",
        vestiging: str = "",
    ) -> str:
        """Genereer een preview zonder API call (dry-run)."""
        return self._assemble_email(
            first_name=first_name,
            company_name=company_name,
            connection_text=connection_placeholder,
            vestiging=vestiging,
        )

    # ── Interne methodes ──────────────────────────────────────────────────

    def _generate_connection_sentences(
        self,
        first_name: str,
        job_title: str,
        company_name: str,
        website: str,
    ) -> str:
        """
        Vraag OpenAI om 2-3 gepersonaliseerde connectiezinnen.
        """
        website_info = f"(website: {website})" if website else "(geen website beschikbaar)"

        prompt = f"""
Zoek op wat {company_name} doet {website_info}.
De ontvanger is {first_name}, {job_title} — hij/zij weet wat zijn/haar bedrijf doet, vat het dus NIET samen.

Schrijf precies 2-3 zinnen (platte tekst, geen opmaak, geen links, geen haakjes) die uitleggen
waarom {self.sender}, student {self.studie} ({self.universiteit}), specifiek bij {company_name} uitkwam.

Schrijf vanuit {self.sender.split()[0]}s perspectief en interesse:
- Welk thema uit zijn studie (logistiek, processen, strategie, data, techniek, operations) 
  is relevant voor hun sector?
- Welk type vraagstuk speelt er in hun branche?
- Wees concreet en specifiek voor dit bedrijf.

Verboden: "innovatief", "onder de indruk", "met interesse gevolgd", URLs, of haakjes met links.

Geef ALLEEN de 2-3 zinnen terug, niets anders.
""".strip()

        response = self.client.responses.create(
            model=self.model,
            input=[{"role": "user", "content": [{"type": "input_text", "text": prompt}]}],
        )

        # Parse output
        output_text = getattr(response, "output_text", None) or ""
        if not output_text:
            for item in (response.output or []):
                if item.type == "message":
                    for c in (item.content or []):
                        if c.type == "output_text":
                            output_text = c.text
                            break

        if not output_text.strip():
            raise RuntimeError("OpenAI gaf een lege response terug")

        return output_text.strip()

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
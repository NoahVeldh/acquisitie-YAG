from __future__ import annotations

import pandas as pd
from src.storage import _normalize_company, _extract_variants

# Laad ruwe DNC lijst
df_dnc = pd.read_excel("data/Niet Benaderen.xlsx", sheet_name="Niet Benaderen")

results = []
for raw in df_dnc["Bedrijf"].dropna():
    raw = str(raw).strip()
    variants = _extract_variants(raw)
    results.append({
        "Originele naam": raw,
        "Genormaliseerd (hoofd)": _normalize_company(raw),
        "Alle varianten": " | ".join(sorted(variants)),
        "Aantal varianten": len(variants),
    })

output_path = "output/dnc_alle_resultaten.xlsx"
df_out = pd.DataFrame(results)
df_out.to_excel(output_path, index=False)

print(f"✅ {len(results)} bedrijven geëxporteerd naar: {output_path}")
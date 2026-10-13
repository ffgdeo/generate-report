# deck-export-demo

Databricks replacement for a Power BI button + Power Automate flow that exported a filtered
PowerPoint deck. Demo data: `fd_fevm_classic_catalog.aibi`.

| Original flow | This demo |
|---|---|
| Power BI button passes selected rows | Streamlit Databricks App (`app/`) |
| Realm `Switch` -> per-realm page list | `REALM_SLIDES` in `src/build_deck.py` (region picks the slide set; mixed -> multi-region) |
| Export API renders PPTX | Serverless job + python-pptx, native editable charts/tables |
| Email attachment | Download button; file also lands in UC Volume `deck_demo.exports` |

Deploy: `databricks bundle deploy && databricks bundle run deck_app` (profile `fevm-fd-classic`),
then grant the app service principal `USE CATALOG` + `USE SCHEMA, SELECT` on the data.
Local test: `.venv/bin/python tests/local_smoke.py`.

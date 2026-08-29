# AGENTS.md

## Project purpose

This repository contains analysis scripts for selected ECG records.

## Project rules

- Treat `/Users/kesju/DI/DUOMENYS/ATRINKTI_ANOTUOTI_IRASAI` as read-only input data.
- Do not modify files in that directory.
- Reuse existing functions from:
  `/Users/kesju/DI/2026_ZIVEO/ECG_ANALYSIS_SANDBOX/ECG_PIPELINE_REPO`
  when appropriate.
- Do not modify `ECG_PIPELINE_REPO` unless explicitly requested.
- Put analysis scripts in `scripts/`.
- Put generated outputs in `results/`.
- Use Python environment `ITP259-ref`.
- Keep analysis-specific logic in this repository.
- Do not copy reusable ECG pipeline functions into this repository unnecessarily.
- ignore manifest.json
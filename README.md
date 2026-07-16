# Formulation Characteristics Pipeline

Agent prompts and pipeline scripts for automated formulation characteristics extraction from PDF labels.

## Structure

- `agents/` — extraction and classification agent prompts (orchestrator, classifiers, extractors)
- `workflow/` — pipeline scripts, validators, and diagnostics utilities

## Setup

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Run

```bash
python workflow/run_pipeline.py
```

#!/usr/bin/env python3
"""Repo-root entry point for the skill's asset renderer.

The implementation lives in skills/pptx/scripts/render_assets.py so it ships
with the Agent Skill; this wrapper lets `python3 scripts/render_assets.py` work
from the repository root as documented in the README.
"""

import runpy
from pathlib import Path

runpy.run_path(
    str(Path(__file__).resolve().parents[1] / "skills/pptx/scripts/render_assets.py"),
    run_name="__main__",
)

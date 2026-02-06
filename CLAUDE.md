# Project Notes & Lessons Learned

## Virtual Environment

- This project uses **uv** for environment management (not pip directly)
- The `.venv` is built on **Python 3.10** (managed by uv, not a system install)
- If the venv breaks, recreate with: `uv venv` then `uv sync`
- Do NOT rely on system Python being available — uv downloads its own Python
- The hardlink warning from `uv sync` is harmless — ignore it

## VS Code Setup

- Always set `python.defaultInterpreterPath` to `${workspaceFolder}/.venv/Scripts/python.exe` in `.vscode/settings.json`
- Set `python-envs.defaultEnvManager` to `ms-python.python:venv` (not `system`) to prevent VS Code from endlessly scanning for environments
- Map the workspace to `.venv` in `python-envs.pythonProjects` to stop the "refreshing virtual environments" loop

## Common Pitfalls

- If Python 3.10 gets uninstalled from `C:\Users\abhis\AppData\Local\Programs\Python\Python310`, the venv will break because `pyvenv.cfg` points to that path — fix by deleting `.venv` and running `uv venv && uv sync`
- Always use `uv sync` (not `pip install -r requirements.txt`) to install deps — it respects the `uv.lock` file for reproducible installs

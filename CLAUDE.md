# Project Notes & Lessons Learned

## Virtual Environment

- This project uses **uv** for environment management (not pip directly)
- The `.venv` is built on **Python 3.10** (managed by uv, not a system install)
- `.python-version` pins 3.10, so `uv venv` reproduces it — do not delete that file.
  Without it uv picks the newest interpreter allowed by `requires-python`, and the
  `numpy<2` pin has no wheels past CPython 3.12
- If the venv breaks, recreate with: `uv venv` then `uv sync`
- Do NOT rely on system Python being available — uv downloads its own Python
- The hardlink warning from `uv sync` is harmless — ignore it
- Prefer `uv run <cmd>` (e.g. `uv run pytest`) over calling `.venv/Scripts/python.exe`
  directly — it checks the environment against `uv.lock` first
- `gunicorn` carries a `sys_platform != 'win32'` marker: it is absent from the local
  venv by design and installs only on the deploy host
- `pyinstaller` lives in the non-default `packaging` group — build releases with
  `uv sync --group packaging`

## Lint, Format and Types

- **ruff** is the only formatter and linter (it replaced flake8, black, isort and
  pylint). **mypy** is the CLI type checker; Pylance covers the editor
- Before committing: `uv run ruff format .`, `uv run ruff check .`, `uv run mypy .`
- All settings live under `[tool.ruff]` / `[tool.mypy]` in `pyproject.toml`. There
  is no `.flake8`, `.pylintrc` or `[tool.black]` — do not recreate them
- Suppress a single finding with `# noqa: <CODE>` plus a reason on the same line.
  RUF100 reports any `noqa` that no longer suppresses anything, so remove those

## VS Code Setup

- Set `python.defaultInterpreterPath` to `${workspaceFolder}/.venv/Scripts/python.exe`
  in `.vscode/settings.json`. That is the only interpreter setting this project needs
- The `python-envs.*` settings were removed. They belong to the separate
  **ms-python.vscode-python-envs** extension, which is not installed here, so they had
  no effect. That extension also has no uv package manager (it ships venv, conda,
  pipenv, poetry and pyenv only), and its `packageManager` defaults to pip — which
  would bypass `uv.lock`. Install packages from a terminal with uv, never from an
  editor package UI

## Common Pitfalls

- If Python 3.10 gets uninstalled from `C:\Users\abhis\AppData\Local\Programs\Python\Python310`, the venv will break because `pyvenv.cfg` points to that path — fix by deleting `.venv` and running `uv venv && uv sync`
- Always use `uv sync` (not `pip install -r requirements.txt`) to install deps — it respects the `uv.lock` file for reproducible installs

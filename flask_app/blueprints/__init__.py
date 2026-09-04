"""
Flask blueprints for the web suite.

A real package rather than an implicit namespace one: without this file the
same module resolves under two names (``excel_gen`` and
``blueprints.excel_gen``) and mypy refuses to type-check the tree.
"""

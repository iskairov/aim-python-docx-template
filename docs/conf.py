# -*- coding: utf-8 -*-
"""
Конфигурационный файл сборки документации Sphinx для проекта python-docx-template.
"""

extensions = [
    "sphinx.ext.autodoc",
]

templates_path = ["_templates"]
source_suffix = ".rst"
master_doc = "index"

project = "python-docx-template"
copyright = "2015-2026, Eric Lapouyade"

version = "0.21"
release = "0.21.0"
language = "ru"

exclude_patterns = ["_build"]
pygments_style = "sphinx"

html_theme = "sphinx_book_theme"
html_static_path = ["_static"]
htmlhelp_basename = "python-docx-templatedoc"

latex_elements = {}
latex_documents = [
    (
        "index",
        "python-docx-template.tex",
        "Документация python-docx-template",
        "Eric Lapouyade",
        "manual",
    ),
]

man_pages = [
    (
        "index",
        "python-docx-template",
        "Документация python-docx-template",
        ["Eric Lapouyade"],
        1,
    )
]

texinfo_documents = [
    (
        "index",
        "python-docx-template",
        "Документация python-docx-template",
        "Eric Lapouyade",
        "python-docx-template",
        "Генерация документов Word (.docx) по шаблонам Jinja2.",
        "Miscellaneous",
    ),
]

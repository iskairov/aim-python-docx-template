# -*- coding: utf-8 -*-
"""
Пакет docxtpl — генерация документов Microsoft Word (.docx) по шаблонам Jinja2.
"""
from __future__ import annotations

__version__ = "0.21.0"

# flake8: noqa
from .inline_image import InlineImage
from .listing import Listing
from .properties import CustomProperties
from .richtext import R, RP, RichText, RichTextParagraph
from .subdoc import Subdoc
from .template import ContextCheckResult, DocxTemplate, TemplateValidationError

__all__ = [
    "__version__",
    "ContextCheckResult",
    "CustomProperties",
    "DocxTemplate",
    "InlineImage",
    "Listing",
    "R",
    "RP",
    "RichText",
    "RichTextParagraph",
    "Subdoc",
    "TemplateValidationError",
]

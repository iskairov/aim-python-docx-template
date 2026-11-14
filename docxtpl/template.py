# -*- coding: utf-8 -*-
"""
Основной модуль шаблонизатора DocxTemplate для генерации документов Word (.docx) с помощью Jinja2.
"""
from __future__ import annotations

import binascii
import functools
from html import escape  # noqa: F401
import io
import os
from os import PathLike
import re
from typing import TYPE_CHECKING, Any, IO
import urllib.parse
import zipfile

from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as REL_TYPE
from docx.opc.oxml import parse_xml
from docx.opc.packuri import PackURI
from docx.opc.part import Part, XmlPart
import docx.oxml.ns
from jinja2 import Environment, meta
from jinja2.exceptions import TemplateError
from jinja2.sandbox import SandboxedEnvironment
from lxml import etree

from .inline_image import InlineImage
from .listing import Listing
from .properties import (
    CORE_STRING_PROPERTIES,
    CustomProperties,
    get_app_properties_dict,
    get_core_properties_dict,
    render_app_properties,
    set_app_properties,
    set_core_property_value,
)
from .richtext import RichText, RichTextParagraph

if TYPE_CHECKING:
    from docx.document import Document as DocxDocument
    from .subdoc import Subdoc


class TemplateValidationError:
    """ Информация об ошибке синтаксиса Jinja2 или структуры XML в части шаблона. """

    def __init__(
        self,
        part_name: str,
        message: str,
        lineno: int | None = None,
        context_lines: list[str] | None = None,
    ) -> None:
        """
        Инициализация информации об ошибке валидации шаблона.

        :param part_name: Название части документа (например, 'body', 'header (rId1)', 'footnotes').
        :param message: Текст сообщения об ошибке синтаксиса Jinja2.
        :param lineno: Номер строки (абзаца) в XML-представлении части документа.
        :param context_lines: Список строк текста соседних абзацев вокруг места ошибки.
        """

        self._part_name = part_name
        self._message = message
        self._lineno = lineno
        self._context_lines = list(context_lines or [])

    @property
    def part_name(self) -> str:
        """ Название части документа, в которой обнаружена ошибка. """

        return self._part_name

    @property
    def message(self) -> str:
        """ Текст описания синтаксической ошибки Jinja2. """

        return self._message

    @property
    def lineno(self) -> int | None:
        """ Номер строки (абзаца) в XML-части документа или None. """

        return self._lineno

    @property
    def context_lines(self) -> list[str]:
        """ Список текстовых строк абзацев вокруг обнаруженной ошибки. """

        return self._context_lines

    def to_dict(self) -> dict[str, Any]:
        """
        Преобразует объект ошибки валидации в словарь.

        :return: Словарь с ключами 'part_name', 'message', 'lineno', 'context_lines'.
        """

        return {
            "part_name": self._part_name,
            "message": self._message,
            "lineno": self._lineno,
            "context_lines": list(self._context_lines),
        }

    def __repr__(self) -> str:
        """
        Возвращает строковое представление ошибки для отладки.

        :return: Строка с частью документа, номером строки и текстом ошибки.
        """

        loc = f"{self._part_name}:{self._lineno}" if self._lineno is not None else self._part_name
        return f"TemplateValidationError({loc!r}, {self._message!r})"

    def __str__(self) -> str:
        """
        Формирует человекочитаемое описание ошибки с контекстом абзацев.

        :return: Отформатированная строка описания ошибки.
        """

        loc = f"{self._part_name} (строка {self._lineno})" if self._lineno is not None else self._part_name
        ctx = "\n  ".join(self._context_lines)
        if ctx:
            return f"[{loc}] {self._message}\n  Контекст:\n  {ctx}"
        return f"[{loc}] {self._message}"


class ContextCheckResult:
    """ Результат проверки полноты контекста и валидности шаблона перед рендерингом. """

    def __init__(
        self,
        required: set[str],
        provided: set[str],
        syntax_errors: list[TemplateValidationError] | None = None,
    ) -> None:
        """
        Инициализация отчёта о проверке контекста шаблона.

        :param required: Множество всех переменных, ожидаемых шаблоном.
        :param provided: Множество ключей, переданных в словаре контекста.
        :param syntax_errors: Список синтаксических ошибок шаблона, если они были найдены.
        """

        self._required = set(required)
        self._provided = set(provided)
        self._missing = self._required - self._provided
        self._unused = self._provided - self._required
        self._syntax_errors = list(syntax_errors or [])

    @property
    def required(self) -> set[str]:
        """ Множество всех переменных, используемых в шаблоне документа. """

        return self._required

    @property
    def provided(self) -> set[str]:
        """ Множество ключей, переданных в проверяемом словаре контекста. """

        return self._provided

    @property
    def missing(self) -> set[str]:
        """ Множество переменных, которые есть в шаблоне, но отсутствуют в контексте. """

        return self._missing

    @property
    def undeclared(self) -> set[str]:
        """ Псевдоним свойства missing (необъявленные переменные шаблона). """

        return self._missing

    @property
    def unused(self) -> set[str]:
        """ Множество ключей контекста, которые нигде не задействованы в шаблоне. """

        return self._unused

    @property
    def syntax_errors(self) -> list[TemplateValidationError]:
        """ Список синтаксических ошибок, обнаруженных в шаблоне. """

        return self._syntax_errors

    @property
    def is_valid(self) -> bool:
        """ Признак полной готовности к рендерингу (нет пропущенных переменных и нет ошибок синтаксиса). """

        return len(self._missing) == 0 and len(self._syntax_errors) == 0

    def to_dict(self) -> dict[str, Any]:
        """
        Возвращает результаты проверки в виде обычного словаря Python.

        :return: Словарь с ключами 'is_valid', 'required', 'provided', 'missing', 'unused', 'syntax_errors'.
        """

        return {
            "is_valid": self.is_valid,
            "required": set(self._required),
            "provided": set(self._provided),
            "missing": set(self._missing),
            "undeclared": set(self._missing),
            "unused": set(self._unused),
            "syntax_errors": list(self._syntax_errors),
        }

    def __getitem__(self, key: str) -> Any:
        """
        Позволяет обращаться к полям результата проверки как к ключам словаря.

        :param key: Имя поля ('is_valid', 'required', 'provided', 'missing', 'undeclared', 'unused', 'syntax_errors').

        :return: Значение соответствующего свойства.
        """

        data = self.to_dict()
        if key not in data:
            raise KeyError(key)
        return data[key]

    def __bool__(self) -> bool:
        """
        Позволяет использовать объект напрямую в логических выражениях if.

        :return: True, если контекст полон и шаблон не содержит синтаксических ошибок.
        """

        return self.is_valid

    def __repr__(self) -> str:
        """
        Возвращает строковое представление результата проверки контекста.

        :return: Строка со сводкой по полям is_valid, missing, unused и syntax_errors.
        """

        return (
            f"ContextCheckResult(is_valid={self.is_valid}, "
            f"missing={sorted(self._missing)}, "
            f"unused={sorted(self._unused)}, "
            f"syntax_errors={len(self._syntax_errors)})"
        )


class DocxTemplate:
    """ Класс для работы с файлами .docx как с шаблонами Jinja2. """

    HEADER_URI = (
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
    )
    FOOTER_URI = (
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
    )
    FOOTNOTES_CONTENT_TYPE = (
        "application/vnd.openxmlformats-officedocument"
        ".wordprocessingml.footnotes+xml"
    )
    ENDNOTES_CONTENT_TYPE = (
        "application/vnd.openxmlformats-officedocument"
        ".wordprocessingml.endnotes+xml"
    )
    COMMENTS_CONTENT_TYPE = (
        "application/vnd.openxmlformats-officedocument"
        ".wordprocessingml.comments+xml"
    )

    def __init__(
        self,
        template_file: IO[bytes] | str | PathLike[str],
        sandbox: bool = False,
    ) -> None:
        """
        Инициализация шаблона документа Word.

        :param template_file: Путь к файлу шаблона .docx или файлоподобный бинарный поток.
        :param sandbox: Использовать ли безопасное песочничное окружение Jinja2 (SandboxedEnvironment) по умолчанию (Issue #644).
        """

        self.template_file = template_file
        self.sandbox = sandbox
        self.reset_replacements()
        self.docx: DocxDocument | None = None
        self._custom_properties: CustomProperties | None = None
        self.is_rendered = False
        self.is_saved = False
        self.allow_missing_pics = False
        self.pic_map: dict[str, Any] = {}
        self.current_rendering_part: Any = None
        self.docx_ids_index = 1000
        self._update_fields_on_open = False

    def init_docx(self, reload: bool = True) -> None:
        """
        Инициализирует внутренний объект python-docx Document из файла шаблона.

        :param reload: Перезагружать ли документ с диска/потока, если он уже был отрендерен.
        """

        if not self.docx or (self.is_rendered and reload):
            if hasattr(self.template_file, "seek"):
                try:
                    self.template_file.seek(0)  # type: ignore[union-attr]
                except Exception:
                    pass
            self.docx = Document(self.template_file)
            self._custom_properties = None
            self.is_rendered = False
            if self._update_fields_on_open:
                self._apply_update_fields_setting(True)

    def render_init(self) -> None:
        """
        Подготавливает внутреннее состояние объекта перед запуском рендеринга.
        """

        self.init_docx()
        self.pic_map = {}
        self.current_rendering_part = None
        self.docx_ids_index = 1000
        self.is_saved = False

    def __getattr__(self, name: str) -> Any:
        """
        Делегирует обращение к неизвестным атрибутам внутреннему объекту python-docx Document.

        :param name: Имя атрибута.

        :return: Значение атрибута объекта self.docx.
        """

        self.init_docx(reload=False)
        return getattr(self.docx, name)

    @property
    def custom_properties(self) -> CustomProperties:
        """ Экземпляр менеджера пользовательских свойств документа Word (docProps/custom.xml). """

        self.init_docx(reload=False)
        if self._custom_properties is None:
            assert self.docx is not None
            self._custom_properties = CustomProperties(self.docx)
        return self._custom_properties

    def get_custom_properties(self) -> dict[str, Any]:
        """
        Возвращает словарь всех пользовательских свойств документа (docProps/custom.xml).

        :return: Словарь {имя_свойства: значение}.
        """

        return self.custom_properties.to_dict()

    def set_custom_property(self, name: str, value: Any) -> None:
        """
        Устанавливает значение одного пользовательского свойства документа Word.

        :param name: Имя пользовательского свойства.
        :param value: Значение свойства (str, int, float, bool, datetime, date).
        """

        self.custom_properties[name] = value

    def set_custom_properties(self, properties: dict[str, Any]) -> None:
        """
        Массово устанавливает или обновляет пользовательские свойства документа Word.

        :param properties: Словарь пользовательских свойств {имя: значение}.
        """

        self.custom_properties.update(properties)

    def delete_custom_property(self, name: str) -> bool:
        """
        Удаляет пользовательское свойство документа по его имени.

        :param name: Имя свойства для удаления.

        :return: True, если свойство существовало и было удалено, иначе False.
        """

        return self.custom_properties.delete(name)

    def get_core_properties(self) -> dict[str, Any]:
        """
        Возвращает словарь основных свойств документа (docProps/core.xml).

        :return: Словарь стандартных свойств (title, author, subject, keywords, comments и др.).
        """

        self.init_docx(reload=False)
        assert self.docx is not None
        return get_core_properties_dict(self.docx)

    def set_core_properties(self, properties: dict[str, Any] | None = None, **kwargs: Any) -> None:
        """
        Устанавливает стандартные свойства документа (docProps/core.xml).

        :param properties: Словарь свойств для установки.
        :param kwargs: Именованные аргументы свойств (например, title='Отчёт', author='Иванов').
        """

        self.init_docx(reload=False)
        assert self.docx is not None
        merged = {**(properties or {}), **kwargs}
        for name, value in merged.items():
            set_core_property_value(self.docx, name, value)

    def get_app_properties(self) -> dict[str, str]:
        """
        Возвращает словарь расширенных свойств приложения из docProps/app.xml.

        :return: Словарь свойств (Company, Manager, Application, AppVersion и др.).
        """

        self.init_docx(reload=False)
        assert self.docx is not None
        return get_app_properties_dict(self.docx)

    def set_app_properties(self, properties: dict[str, Any] | None = None, **kwargs: Any) -> None:
        """
        Устанавливает расширенные свойства документа в docProps/app.xml (например, Company, Manager).

        :param properties: Словарь устанавливаемых свойств.
        :param kwargs: Именованные параметры свойств.
        """

        self.init_docx(reload=False)
        assert self.docx is not None
        merged = {**(properties or {}), **kwargs}
        set_app_properties(self.docx, merged)

    def get_all_properties(self) -> dict[str, dict[str, Any]]:
        """
        Возвращает все свойства документа, сгруппированные по категориям (core, app, custom).

        :return: Словарь с ключами 'core', 'app' и 'custom'.
        """

        return {
            "core": self.get_core_properties(),
            "app": self.get_app_properties(),
            "custom": self.get_custom_properties(),
        }

    def _apply_update_fields_setting(self, enabled: bool) -> None:
        """
        Добавляет или удаляет элемент <w:updateFields w:val="true"/> в настройках документа (word/settings.xml).

        :param enabled: Включить (True) или отключить (False) автообновление полей при открытии в Word.
        """

        if self.docx is None:
            return
        settings_el = self.docx.settings.element
        w_ns = docx.oxml.ns.nsmap["w"]
        tag = f"{{{w_ns}}}updateFields"
        val_attr = f"{{{w_ns}}}val"
        existing = settings_el.find(tag)

        if enabled:
            if existing is None:
                existing = etree.SubElement(settings_el, tag)
            existing.set(val_attr, "true")
        else:
            if existing is not None:
                settings_el.remove(existing)

    def update_fields_on_open(self, enabled: bool = True) -> None:
        """
        Управляет автоматическим пересчётом всех полей документа (оглавления TOC, нумерации страниц,
        перекрёстных ссылок REF, вычисляемых формул таблиц и DOCPROPERTY) при открытии файла в Microsoft Word.

        :param enabled: Если True — прописывает <w:updateFields w:val="true"/> в word/settings.xml.
        """

        self._update_fields_on_open = enabled
        self.init_docx(reload=False)
        self._apply_update_fields_setting(enabled)

    def get_update_fields_on_open(self) -> bool:
        """
        Проверяет, включён ли автоматический пересчёт полей (<w:updateFields>) в настройках документа.

        :return: True, если автообновление полей при открытии активно, иначе False.
        """

        self.init_docx(reload=False)
        assert self.docx is not None
        settings_el = self.docx.settings.element
        w_ns = docx.oxml.ns.nsmap["w"]
        existing = settings_el.find(f"{{{w_ns}}}updateFields")
        if existing is None:
            return False
        val = existing.get(f"{{{w_ns}}}val", "true").strip().lower()
        return val in ("true", "1", "on")

    def xml_to_string(self, xml: etree._Element, encoding: str = "unicode") -> str:
        """
        Сериализует XML-элемент lxml в строку без форматирования отступами.

        :param xml: Корневой XML-элемент для сериализации.
        :param encoding: Кодировка результата (по умолчанию 'unicode').

        :return: Строковое представление XML.
        """

        return etree.tostring(xml, encoding="unicode", pretty_print=False)

    def get_docx(self) -> DocxDocument:
        """
        Инициализирует (при необходимости) и возвращает внутренний объект python-docx Document.

        :return: Экземпляр документа python-docx.
        """

        self.init_docx()
        assert self.docx is not None
        return self.docx

    def get_xml(self) -> str:
        """
        Возвращает XML-строку тела документа (<w:body>).

        :return: Строка XML тела документа.
        """

        self.init_docx(reload=False)
        assert self.docx is not None
        return self.xml_to_string(self.docx._element.body)

    def write_xml(self, filename: str | PathLike[str]) -> None:
        """
        Записывает текущее XML-тело документа в текстовый файл в кодировке UTF-8.

        :param filename: Путь к файлу для сохранения XML.
        """

        with open(filename, "w", encoding="utf-8") as fh:
            fh.write(self.get_xml())

    @staticmethod
    def _normalize_math_xml(src_xml: str) -> str:
        """
        Нормализует структуру математических формул OMML (<m:oMath>) для корректной замены тегов Jinja2.

        :param src_xml: Исходная XML-строка части документа.

        :return: XML-строка с нормализованными математическими блоками <m:oMath>.
        """

        if "<m:oMath" not in src_xml:
            return src_xml

        def process_omath(match: re.Match[str]) -> str:
            omath_xml = match.group(0)

            # 1. Разворачиваем элементы разделителей <m:d>, если они представляют фигурные скобки '{' и '}'
            d_pattern = re.compile(
                r"<m:d>(?:\s*<m:dPr>(?:(?!</m:dPr>).)*?<m:begChr\s+m:val=\"\{\"\s*/>"
                r"(?:(?!</m:dPr>).)*?<m:endChr\s+m:val=\"\}\"\s*/>(?:(?!</m:dPr>).)*?</m:dPr>)"
                r"\s*<m:e>((?:(?!<m:d>).)*?)</m:e>\s*</m:d>",
                flags=re.DOTALL,
            )
            for _ in range(4):
                new_xml = d_pattern.sub(
                    r"<m:r><m:t>{</m:t></m:r>\1<m:r><m:t>}</m:t></m:r>",
                    omath_xml,
                )
                if new_xml == omath_xml:
                    break
                omath_xml = new_xml

            # 2. Склеиваем разорванные соседние скобки { и {, } и } внутри <m:t>
            for _ in range(3):
                omath_xml = re.sub(
                    r"(?<={)(?:<(?!/m:oMath>)[^>]*>)+(?=[\{%\#])|(?<=[%\}\#])(?:<(?!/m:oMath>)[^>]*>)+(?=\})",
                    "",
                    omath_xml,
                    flags=re.DOTALL,
                )

            # 3. Если внутри тега Jinja2 редактор формул превратил 'my_var' в нижний индекс <m:sSub>,
            # восстанавливаем плоскую запись 'my_var'
            def fix_ssub_in_jinja(m_tag: re.Match[str]) -> str:
                tag_block = m_tag.group(0)
                if "<m:sSub" not in tag_block:
                    return tag_block

                def flatten_ssub(m_sub: re.Match[str]) -> str:
                    e_content = re.sub(r"<[^>]+>", "", m_sub.group(1))
                    sub_content = re.sub(r"<[^>]+>", "", m_sub.group(2))
                    return f"<m:r><m:t>{e_content}_{sub_content}</m:t></m:r>"

                return re.sub(
                    r"<m:sSub>(?:\s*<m:sSubPr>.*?</m:sSubPr>)?\s*<m:e>(.*?)</m:e>\s*<m:sub>(.*?)</m:sub>\s*</m:sSub>",
                    flatten_ssub,
                    tag_block,
                    flags=re.DOTALL,
                )

            omath_xml = re.sub(
                r"\{%(?:(?!%\}).)*%\}|\{\{(?:(?!\}\}).)*\}\}|\{#(?:(?!#\}).)*#\}",
                fix_ssub_in_jinja,
                omath_xml,
                flags=re.DOTALL,
            )

            # Случай, когда открывающая '{{' оказалась внутри <m:e>, а закрывающая '}}' — внутри <m:sub> одного <m:sSub>
            def fix_split_ssub(m_sub: re.Match[str]) -> str:
                e_part = m_sub.group(1)
                sub_part = m_sub.group(2)
                e_text = re.sub(r"<[^>]+>", "", e_part)
                sub_text = re.sub(r"<[^>]+>", "", sub_part)
                has_unclosed_open = (
                    ("{{" in e_text and "}}" not in e_text.split("{{")[-1])
                    or ("{%" in e_text and "%}" not in e_text.split("{%")[-1])
                )
                has_close = ("}}" in sub_text) or ("%}" in sub_text)
                if has_unclosed_open and has_close:
                    return f"<m:r><m:t>{e_text}_{sub_text}</m:t></m:r>"
                return m_sub.group(0)

            omath_xml = re.sub(
                r"<m:sSub>(?:\s*<m:sSubPr>.*?</m:sSubPr>)?\s*<m:e>(.*?)</m:e>\s*<m:sub>(.*?)</m:sub>\s*</m:sSub>",
                fix_split_ssub,
                omath_xml,
                flags=re.DOTALL,
            )

            # 4. Склеиваем разорванные <m:t>...</m:t> внутри каждого тега Jinja2 в формуле
            def strip_math_tags(m_tag: re.Match[str]) -> str:
                return re.sub(
                    r"</m:t>(?:(?!</m:oMath>|</m:e>|</m:num>|</m:den>|</m:sub>|</m:sup>).)*?<m:t(?:\s[^>]*)?>",
                    "",
                    m_tag.group(0),
                    flags=re.DOTALL,
                )

            omath_xml = re.sub(
                r"\{%(?:(?!%\}).)*%\}|\{\{(?:(?!\}\}).)*\}\}|\{#(?:(?!#\}).)*#\}",
                strip_math_tags,
                omath_xml,
                flags=re.DOTALL,
            )

            return omath_xml

        return re.sub(
            r"<m:oMath[ >].*?</m:oMath>",
            process_omath,
            src_xml,
            flags=re.DOTALL,
        )

    def patch_xml(self, src_xml: str) -> str:
        """
        Очищает и подготавливает сырой XML-код документа для корректной обработки шаблонизатором Jinja2.

        :param src_xml: Исходная XML-строка части документа.

        :return: Очищенная и подготовленная для Jinja2 XML-строка.
        """

        # Предварительная нормализация математических формул OMML (<m:oMath>)
        src_xml = self._normalize_math_xml(src_xml)

        # Склеиваем разорванные разделители вида {<теги>{ в {{ (работает для {{ }}, {% %}, {# #})
        for _ in range(3):
            src_xml = re.sub(
                r"(?<={)(?:[\u200b\ufeff]|<(?!/w:p>|/w:tc>|/w:tr>|/m:oMath>)[^>]*>)+(?=[\{%\#])"
                r"|(?<=[%\}\#])(?:[\u200b\ufeff]|<(?!/w:p>|/w:tc>|/w:tr>|/m:oMath>)[^>]*>)+(?=\})",
                "",
                src_xml,
                flags=re.DOTALL,
            )

        # Внутри выражений {{ ... }}, {% ... %} и {# ... #} удаляем промежуточные теги Word, формул, фигур и полей
        def striptags(m: re.Match[str]) -> str:
            block = m.group(0)
            block = re.sub(
                r"</w:t>.*?(<w:t>|<w:t [^>]*>)",
                "",
                block,
                flags=re.DOTALL,
            )
            block = re.sub(
                r"</m:t>.*?(<m:t>|<m:t [^>]*>)",
                "",
                block,
                flags=re.DOTALL,
            )
            block = re.sub(
                r"</a:t>.*?(<a:t>|<a:t [^>]*>)",
                "",
                block,
                flags=re.DOTALL,
            )
            block = re.sub(
                r"</w:instrText>.*?(<w:instrText>|<w:instrText [^>]*>)",
                "",
                block,
                flags=re.DOTALL,
            )
            return block

        src_xml = re.sub(
            r"{%(?:(?!%}).)*|{#(?:(?!#}).)*|{{(?:(?!}}).)*",
            striptags,
            src_xml,
            flags=re.DOTALL,
        )

        # Обработка горизонтального объединения ячеек таблицы: {% colspan expr %}
        def colspan(m: re.Match[str]) -> str:
            cell_xml = m.group(1) + m.group(3)
            cell_xml = re.sub(
                r"<w:r[ >](?:(?!<w:r[ >]).)*<w:t></w:t>.*?</w:r>",
                "",
                cell_xml,
                flags=re.DOTALL,
            )
            cell_xml = re.sub(r"<w:gridSpan[^/]*/>", "", cell_xml, count=1)
            return re.sub(
                r"(<w:tcPr[^>]*>)",
                r'\1<w:gridSpan w:val="{{%s}}"/>' % m.group(2).strip(),
                cell_xml,
            )

        src_xml = re.sub(
            r"(<w:tc[ >](?:(?!<w:tc[ >]).)*){%\s*colspan\s+((?:(?!%}).)+?)\s*%}(.*?</w:tc>)",
            colspan,
            src_xml,
            flags=re.DOTALL,
        )

        # Обработка цвета фона ячейки таблицы: {% cellbg expr %}
        def cellbg(m: re.Match[str]) -> str:
            cell_xml = m.group(1) + m.group(3)
            cell_xml = re.sub(
                r"<w:r[ >](?:(?!<w:r[ >]).)*<w:t></w:t>.*?</w:r>",
                "",
                cell_xml,
                flags=re.DOTALL,
            )
            cell_xml = re.sub(r"<w:shd[^/]*/>", "", cell_xml, count=1)
            return re.sub(
                r"(<w:tcPr[^>]*>)",
                r'\1<w:shd w:val="clear" w:color="auto" w:fill="{{%s}}"/>'
                % m.group(2).strip(),
                cell_xml,
            )

        src_xml = re.sub(
            r"(<w:tc[ >](?:(?!<w:tc[ >]).)*){%\s*cellbg\s+((?:(?!%}).)+?)\s*%}(.*?</w:tc>)",
            cellbg,
            src_xml,
            flags=re.DOTALL,
        )

        # Гарантируем сохранение пробелов в текстовых узлах с тегами Jinja2
        src_xml = re.sub(
            r"<w:t>((?:(?!<w:t>|</w:t>).)*)({{.*?}}|{%.*?%})",
            r'<w:t xml:space="preserve">\1\2',
            src_xml,
            flags=re.DOTALL,
        )
        src_xml = re.sub(
            r"({{r\s(?:(?!}}).)*}}|{%r\s(?:(?!%}).)*%})",
            r'</w:t></w:r><w:r><w:t xml:space="preserve">\1</w:t></w:r><w:r><w:t xml:space="preserve">',
            src_xml,
            flags=re.DOTALL,
        )

        # {%- склеивает тег с текстом предыдущего абзаца
        src_xml = re.sub(r"</w:t>(?:(?!</w:t>).)*?{%-", "{%", src_xml, flags=re.DOTALL)
        # -%} склеивает тег с текстом следующего абзаца
        src_xml = re.sub(
            r"-%}(?:(?!<w:t[ >]|{%|{{).)*?<w:t[^>]*?>", "%}", src_xml, flags=re.DOTALL
        )

        # Заменяем строки таблицы (tr), ячейки (tc), абзацы (p) и прогоны (r), содержащие {%y xxx %} или {{y xxx}} (Issue #660)
        for y in ["tr", "tc", "p", "r"]:
            pat = (
                r"<w:%(y)s[ >](?:(?!<w:%(y)s[ >]).)*({%%|{{)%(y)s ((?:(?!%%}|}}).)*(?:%%}|}}))(?:(?!</w:%(y)s>).)*?</w:%(y)s>"
                % {"y": y}
            )
            src_xml = re.sub(pat, r"\1 \2", src_xml, flags=re.DOTALL)

        # Поддержка для математических прогонов <m:r> при использовании {%m ... %} или {{m ... }}
        pat_math_r = (
            r"<m:r[ >](?:(?!<m:r[ >]).)*({%|{{)m ((?:(?!%}|}}).)*(?:%}|}}))(?:(?!</m:r>).)*?</m:r>"
        )
        src_xml = re.sub(pat_math_r, r"\1 \2", src_xml, flags=re.DOTALL)

        # Обработка комментариев {#p ... #}, {#tr ... #}, {#tc ... #}
        for y in ["tr", "tc", "p"]:
            pat = (
                r"<w:%(y)s[ >](?:(?!<w:%(y)s[ >]).)*({#)%(y)s ((?:(?!#}).)*#})(?:(?!</w:%(y)s>).)*?</w:%(y)s>"
                % {"y": y}
            )
            src_xml = re.sub(pat, r"\1 \2", src_xml, flags=re.DOTALL)

        # Вертикальное объединение ячеек таблицы в цикле: {% vm %} или {% vm условие %} (Issue #620)
        def v_merge_tc(m: re.Match[str]) -> str:
            tc_xml = m.group(0)
            vm_match = re.search(r"{%\s*vm(?:\s+((?:(?!%}).)+?))?\s*%}", tc_xml, flags=re.DOTALL)
            cond_expr = (vm_match.group(1).strip() if (vm_match and vm_match.group(1)) else "loop.first")

            if "</w:tcPr" not in tc_xml:
                tc_xml = re.sub(r"^(\s*<w:tc[^>]*>)", r"\1<w:tcPr></w:tcPr>", tc_xml, count=1)

            def v_merge(m1: re.Match[str]) -> str:
                return (
                    f'<w:vMerge w:val="{{% if {cond_expr} %}}restart{{% else %}}continue{{% endif %}}"/>'
                    + m1.group(1)
                    + f"{{% if {cond_expr} %}}"
                    + m1.group(2)
                    + m1.group(3)
                    + "{% endif %}"
                    + m1.group(4)
                )

            return re.sub(
                r"(</w:tcPr[ >].*?<w:t(?:.*?)>)(.*?)(?:{%\s*vm(?:\s+(?:(?!%}).)+?)?\s*%})(.*?)(</w:t>)",
                v_merge,
                tc_xml,
                flags=re.DOTALL,
            )

        src_xml = re.sub(
            r"<w:tc[ >](?:(?!<w:tc[ >]).)*?{%\s*vm(?:\s+(?:(?!%}).)+?)?\s*%}.*?</w:tc[ >]",
            v_merge_tc,
            src_xml,
            flags=re.DOTALL,
        )

        # Горизонтальное объединение ячеек таблицы в цикле: {% hm %}
        def h_merge_tc(m: re.Match[str]) -> str:
            xml_to_patch = m.group()

            def with_gridspan(m1: re.Match[str]) -> str:
                return (
                    m1.group(1)
                    + "{{ "
                    + m1.group(2)
                    + " * loop.length }}"
                    + m1.group(3)
                )

            def without_gridspan(m2: re.Match[str]) -> str:
                return (
                    '<w:gridSpan w:val="{{ loop.length }}"/>'
                    + m2.group(1)
                    + m2.group(2)
                    + m2.group(3)
                    + m2.group(4)
                )

            if re.search(r"w:gridSpan", xml_to_patch):
                xml = re.sub(
                    r'(w:gridSpan w:val=")(\d+)(")',
                    with_gridspan,
                    xml_to_patch,
                    flags=re.DOTALL,
                )
                xml = re.sub(
                    r"{%\s*hm\s*%}",
                    "",
                    xml,
                    flags=re.DOTALL,
                )
            else:
                if "</w:tcPr" not in xml_to_patch:
                    xml_to_patch = re.sub(
                        r"^(\s*<w:tc[^>]*>)",
                        r"\1<w:tcPr></w:tcPr>",
                        xml_to_patch,
                        count=1,
                    )
                xml = re.sub(
                    r"(</w:tcPr[ >].*?<w:t(?:.*?)>)(.*?)(?:{%\s*hm\s*%})(.*?)(</w:t>)",
                    without_gridspan,
                    xml_to_patch,
                    flags=re.DOTALL,
                )

            return "{% if loop.first %}" + xml + "{% endif %}"

        src_xml = re.sub(
            r"<w:tc[ >](?:(?!<w:tc[ >]).)*?{%\s*hm\s*%}.*?</w:tc[ >]",
            h_merge_tc,
            src_xml,
            flags=re.DOTALL,
        )

        # Очистка содержимого тегов Jinja2 от типографских кавычек Word, HTML-сущностей и неразрывных пробелов (Issue #628, #660)
        def clean_tags(m: re.Match[str]) -> str:
            full = m.group(0)
            prefix = full[:2]
            suffix = full[-2:]
            inner = full[2:-2]
            inner = (
                inner.replace(r"&#8216;", "'")
                .replace(r"&#8217;", "'")
                .replace(r"&#8220;", '"')
                .replace(r"&#8221;", '"')
                .replace("&lt;", "<")
                .replace("&gt;", ">")
                .replace("&quot;", '"')
                .replace("&apos;", "'")
                .replace("“", '"')
                .replace("”", '"')
                .replace("‘", "'")
                .replace("’", "'")
                .replace("\u202f", " ")
                .replace("\u00a0", " ")
                .replace("\u200b", "")
                .replace("\ufeff", "")
            )
            return prefix + inner + suffix

        src_xml = re.sub(
            r"\{%(?:(?!%\}).)*%\}|\{\{(?:(?!\}\}).)*\}\}",
            clean_tags,
            src_xml,
            flags=re.DOTALL,
        )

        return src_xml

    def _register_builtin_filters(self, jinja_env: Environment) -> None:
        """
        Регистрирует встроенные фильтры docxtpl (richtext, listing, inline_image) в окружении Jinja2 (Issue #582).

        :param jinja_env: Окружение Jinja2, в которое добавляются фильтры.
        """

        if "richtext" not in jinja_env.filters:
            jinja_env.filters["richtext"] = lambda val, **kw: RichText(val, **kw)
        if "r" not in jinja_env.filters:
            jinja_env.filters["r"] = lambda val, **kw: RichText(val, **kw)
        if "listing" not in jinja_env.filters:
            jinja_env.filters["listing"] = lambda val: Listing(val)
        if "inline_image" not in jinja_env.filters:
            jinja_env.filters["inline_image"] = (
                lambda val, width=None, height=None, **kw: InlineImage(
                    self, val, width=width, height=height, **kw
                )
            )

    def _get_or_create_jinja_env(
        self,
        jinja_env: Environment | None = None,
        autoescape: bool = False,
        sandbox: bool | None = None,
    ) -> Environment:
        """
        Возвращает подготовленное окружение Jinja2 с учётом настроек autoescape и песочницы (sandbox).

        :param jinja_env: Пользовательское окружение Jinja2 или None.
        :param autoescape: Включить ли автоматическое экранирование XML/HTML.
        :param sandbox: Использовать ли SandboxedEnvironment (если None — берётся self.sandbox).

        :return: Экземпляр окружения Jinja2.
        """

        use_sandbox = self.sandbox if sandbox is None else sandbox
        if jinja_env is None:
            env_cls = SandboxedEnvironment if use_sandbox else Environment
            jinja_env = env_cls(autoescape=autoescape)
        elif autoescape:
            jinja_env.autoescape = autoescape

        self._register_builtin_filters(jinja_env)
        return jinja_env

    def render_xml_part(
        self,
        src_xml: str,
        part: Any,
        context: dict[str, Any],
        jinja_env: Environment | None = None,
    ) -> str:
        """
        Рендерит подготовленную XML-строку конкретной части документа с помощью Jinja2.

        :param src_xml: Подготовленная методом patch_xml() XML-строка.
        :param part: Часть пакета документа (используется для привязки изображений и ссылок).
        :param context: Словарь переменных контекста для подстановки.
        :param jinja_env: Окружение Jinja2 или None.

        :return: Отрендеренная XML-строка.
        """

        src_xml = re.sub(r"<w:p([ >])", r"\n<w:p\1", src_xml)
        try:
            self.current_rendering_part = part
            if jinja_env is not None:
                self._register_builtin_filters(jinja_env)
                template = jinja_env.from_string(src_xml)
            elif self.sandbox:
                env = self._get_or_create_jinja_env(sandbox=True)
                template = env.from_string(src_xml)
            else:
                env = self._get_or_create_jinja_env(sandbox=False)
                template = env.from_string(src_xml)
            dst_xml = template.render(context)
        except TemplateError as exc:
            if hasattr(exc, "lineno") and exc.lineno is not None:
                line_number = max(exc.lineno - 4, 0)
                exc.docx_context = map(  # type: ignore[attr-defined]
                    lambda x: re.sub(r"<[^>]+>", "", x),
                    src_xml.splitlines()[line_number : (line_number + 7)],
                )

            raise exc
        dst_xml = re.sub(r"\n<w:p([ >])", r"<w:p\1", dst_xml)
        dst_xml = (
            dst_xml.replace("{_{", "{{")
            .replace("}_}", "}}")
            .replace("{_%", "{%")
            .replace("%_}", "%}")
        )
        dst_xml = self.resolve_listing(dst_xml)
        return dst_xml

    def render_properties(
        self,
        context: dict[str, Any],
        jinja_env: Environment | None = None,
    ) -> None:
        """
        Рендерит шаблонные переменные Jinja2 в стандартных (Core), расширенных (App) и пользовательских (Custom) свойствах документа.

        :param context: Словарь контекста шаблона.
        :param jinja_env: Окружение Jinja2 или None.
        """

        assert self.docx is not None
        env = self._get_or_create_jinja_env(jinja_env)

        # 1. Рендеринг всех строковых стандартных свойств (docProps/core.xml)
        for prop in CORE_STRING_PROPERTIES:
            initial = getattr(self.docx.core_properties, prop, "")
            if initial and ("{{" in initial or "{%" in initial or "{#" in initial):
                template = env.from_string(initial)
                rendered = template.render(context)
                set_core_property_value(self.docx, prop, rendered)

        # 2. Рендеринг расширенных свойств приложения (docProps/app.xml)
        render_app_properties(self.docx, context, env)

        # 3. Рендеринг пользовательских свойств Word (docProps/custom.xml)
        self.custom_properties.render_properties(context, env)

    def render_footnotes(
        self,
        context: dict[str, Any],
        jinja_env: Environment | None = None,
    ) -> None:
        """
        Рендерит шаблонные переменные Jinja2 в сносках (footnotes), концевых сносках (endnotes) и примечаниях (comments).

        :param context: Словарь контекста шаблона.
        :param jinja_env: Окружение Jinja2 или None.
        """

        assert self.docx is not None
        env = self._get_or_create_jinja_env(jinja_env)

        package = self.docx.part.package
        if package is None:
            return

        target_content_types = {
            self.FOOTNOTES_CONTENT_TYPE,
            self.ENDNOTES_CONTENT_TYPE,
            self.COMMENTS_CONTENT_TYPE,
        }

        visited_parts: set[int] = set()
        for part in package.parts:
            if id(part) in visited_parts:
                continue
            visited_parts.add(id(part))

            if part.content_type in target_content_types and part.blob:
                raw_xml = (
                    part.blob.decode("utf-8")
                    if isinstance(part.blob, bytes)
                    else part.blob
                )
                xml = self.patch_xml(raw_xml)
                xml = self.render_xml_part(xml, part, context, env)
                encoded = xml.encode("utf-8")
                part._blob = encoded
                if isinstance(part, XmlPart) and hasattr(part, "_element"):
                    try:
                        part._element = parse_xml(encoded)
                    except Exception:
                        pass

    def resolve_listing(self, xml: str) -> str:
        """
        Преобразует спецсимволы (\t, \a, \n, \f) внутри текстовых узлов <w:t> в соответствующие теги OpenXML.

        :param xml: XML-строка после рендеринга Jinja2.

        :return: XML-строка с преобразованными управляющими символами.
        """

        def resolve_text(
            run_properties: str,
            paragraph_properties: str,
            m: re.Match[str],
        ) -> str:
            text_xml = m.group(0).replace(
                "\t",
                "</w:t></w:r>"
                f"<w:r>{run_properties}<w:tab/></w:r>"
                f'<w:r>{run_properties}<w:t xml:space="preserve">',
            )
            text_xml = text_xml.replace(
                "\a",
                "</w:t></w:r></w:p>"
                f'<w:p>{paragraph_properties}<w:r>{run_properties}<w:t xml:space="preserve">',
            )
            text_xml = text_xml.replace("\n", '</w:t><w:br/><w:t xml:space="preserve">')
            text_xml = text_xml.replace(
                "\f",
                "</w:t></w:r></w:p>"
                '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
                f'<w:p>{paragraph_properties}<w:r>{run_properties}<w:t xml:space="preserve">',
            )
            return text_xml

        def resolve_run(paragraph_properties: str, m: re.Match[str]) -> str:
            run_properties_match = re.search(r"<w:rPr>.*?</w:rPr>", m.group(0))
            run_properties = (
                run_properties_match.group(0) if run_properties_match else ""
            )
            return re.sub(
                r"<w:t(?: [^>]*)?>.*?</w:t>",
                lambda x: resolve_text(run_properties, paragraph_properties, x),
                m.group(0),
                flags=re.DOTALL,
            )

        def resolve_paragraph(m: re.Match[str]) -> str:
            paragraph_properties_match = re.search(r"<w:pPr>.*?</w:pPr>", m.group(0))
            paragraph_properties = (
                paragraph_properties_match.group(0)
                if paragraph_properties_match
                else ""
            )
            return re.sub(
                r"<w:r(?: [^>]*)?>.*?</w:r>",
                lambda x: resolve_run(paragraph_properties, x),
                m.group(0),
                flags=re.DOTALL,
            )

        xml = re.sub(
            r"<w:p(?: [^>]*)?>.*?</w:p>", resolve_paragraph, xml, flags=re.DOTALL
        )

        return xml

    def build_xml(
        self,
        context: dict[str, Any],
        jinja_env: Environment | None = None,
    ) -> str:
        """
        Извлекает, подготавливает и рендерит XML основного тела документа.

        :param context: Словарь переменных контекста.
        :param jinja_env: Окружение Jinja2 или None.

        :return: Отрендеренная XML-строка тела документа.
        """

        assert self.docx is not None
        xml = self.get_xml()
        xml = self.patch_xml(xml)
        xml = self.render_xml_part(xml, self.docx._part, context, jinja_env)
        return xml

    def map_tree(self, tree: etree._Element) -> None:
        """
        Заменяет текущее тело документа (<w:body>) новым XML-деревом.

        :param tree: Новое XML-дерево элемента <w:body>.
        """

        assert self.docx is not None
        root = self.docx._element
        body = root.body
        root.replace(body, tree)

    def get_headers_footers(self, uri: str) -> Any:
        """
        Итерируется по всем частям колонтитулов заданного типа связи (верхние или нижние).

        :param uri: URI типа связи (HEADER_URI или FOOTER_URI).

        :return: Генератор кортежей (relKey, target_part).
        """

        assert self.docx is not None
        for rel_key, val in self.docx._part.rels.items():
            if (val.reltype == uri) and (val.target_part.blob):
                yield rel_key, val.target_part

    def get_part_xml(self, part: Any) -> str:
        """
        Извлекает XML-строку из бинарного содержимого указанной части пакета.

        :param part: Часть пакета документа.

        :return: Строковое представление XML.
        """

        return self.xml_to_string(parse_xml(part.blob))

    def get_headers_footers_encoding(self, xml: str) -> str:
        """
        Определяет кодировку XML-колонтитула по заголовку объявления XML (по умолчанию utf-8).

        :param xml: XML-строка колонтитула.

        :return: Название кодировки.
        """

        m = re.match(r'<\?xml[^\?]+\bencoding="([^"]+)"', xml, re.I)
        if m:
            return m.group(1)
        return "utf-8"

    def build_headers_footers_xml(
        self,
        context: dict[str, Any],
        uri: str,
        jinja_env: Environment | None = None,
    ) -> Any:
        """
        Подготавливает и рендерит XML для всех колонтитулов указанного типа.

        :param context: Словарь переменных контекста.
        :param uri: URI типа связи колонтитула (HEADER_URI или FOOTER_URI).
        :param jinja_env: Окружение Jinja2 или None.

        :return: Генератор кортежей (relKey, закодированный_xml_в_байтах).
        """

        for rel_key, part in self.get_headers_footers(uri):
            xml = self.get_part_xml(part)
            encoding = self.get_headers_footers_encoding(xml)
            xml = self.patch_xml(xml)
            xml = self.render_xml_part(xml, part, context, jinja_env)
            yield rel_key, xml.encode(encoding)

    def map_headers_footers_xml(self, relKey: str, xml: bytes) -> None:
        """
        Заменяет содержимое части колонтитула отрендеренным XML с сохранением всех вложенных связей.

        :param relKey: Ключ связи (rId) колонтитула в основном документе.
        :param xml: Байтовое представление отрендеренного XML колонтитула.
        """

        assert self.docx is not None
        part = self.docx._part.rels[relKey].target_part
        new_part = XmlPart.load(part.partname, part.content_type, xml, part.package)
        for r_id, rel in part.rels.items():
            new_part.load_rel(rel.reltype, rel._target, rel.rId, rel.is_external)
        self.docx._part.rels[relKey]._target = new_part

    def update_docproperty_fields(self) -> None:
        """
        Обновляет кешированные значения всех полей DOCPROPERTY в теле документа и колонтитулах
        актуальными значениями из пользовательских (Custom), основных (Core) и расширенных (App) свойств.
        """

        self.init_docx(reload=False)
        assert self.docx is not None

        extra_props: dict[str, Any] = {}
        extra_props.update(self.get_core_properties())
        extra_props.update(self.get_app_properties())

        cprops = self.custom_properties
        cprops.update_fields_in_element(self.docx._element.body, extra_properties=extra_props)

        for uri in (self.HEADER_URI, self.FOOTER_URI):
            for rel_key, part in self.get_headers_footers(uri):
                if isinstance(part, XmlPart) and hasattr(part, "_element") and part._element is not None:
                    cprops.update_fields_in_element(part._element, extra_properties=extra_props)
                elif part.blob:
                    el = parse_xml(part.blob)
                    cprops.update_fields_in_element(el, extra_properties=extra_props)
                    self.map_headers_footers_xml(rel_key, self.xml_to_string(el).encode("utf-8"))

    def render(
        self,
        context: dict[str, Any],
        jinja_env: Environment | None = None,
        autoescape: bool = False,
        sandbox: bool | None = None,
        update_fields: bool = False,
    ) -> None:
        """
        Выполняет полный рендеринг шаблона документа (тело, таблицы, колонтитулы, свойства, сноски).

        :param context: Словарь переменных контекста для подстановки в шаблон.
        :param jinja_env: Опциональное пользовательское окружение Jinja2.
        :param autoescape: Автоматически экранировать спецсимволы XML в переменных контекста.
        :param sandbox: Использовать ли песочницу SandboxedEnvironment (если None — используется self.sandbox).
        :param update_fields: Включить ли автоматический пересчёт полей (TOC, PAGE, REF, DOCPROPERTY) при открытии документа в Word.
        """

        self.render_init()
        if update_fields or self._update_fields_on_open:
            self._apply_update_fields_setting(True)

        env = self._get_or_create_jinja_env(
            jinja_env=jinja_env,
            autoescape=autoescape,
            sandbox=sandbox,
        )

        # Рендерим основное тело документа
        xml_src = self.build_xml(context, env)

        # Корректируем сетку таблиц при динамическом изменении колонок
        tree = self.fix_tables(xml_src)

        # Перенумеровываем идентификаторы объектов и проверяем встроенные OLE-объекты
        self.fix_docpr_ids(tree)

        # Заменяем дерево <w:body> в документе
        self.map_tree(tree)

        # Рендерим верхние колонтитулы (Headers)
        headers = self.build_headers_footers_xml(context, self.HEADER_URI, env)
        for rel_key, xml in headers:
            self.map_headers_footers_xml(rel_key, xml)

        # Рендерим нижние колонтитулы (Footers)
        footers = self.build_headers_footers_xml(context, self.FOOTER_URI, env)
        for rel_key, xml in footers:
            self.map_headers_footers_xml(rel_key, xml)

        # Рендерим свойства документа (Core, App и Custom)
        self.render_properties(context, env)

        # Обновляем кешированные поля DOCPROPERTY актуальными значениями свойств
        self.update_docproperty_fields()

        # Рендерим сноски, концевые сноски и примечания
        self.render_footnotes(context, env)

        self.is_rendered = True

    async def render_async(
        self,
        context: dict[str, Any],
        jinja_env: Environment | None = None,
        autoescape: bool = False,
        sandbox: bool | None = None,
        update_fields: bool = False,
    ) -> None:
        """
        Асинхронный метод для рендеринга шаблона в корутинах asyncio (Issue #611).

        :param context: Словарь переменных контекста для подстановки в шаблон.
        :param jinja_env: Опциональное пользовательское окружение Jinja2.
        :param autoescape: Автоматически экранировать спецсимволы XML в переменных контекста.
        :param sandbox: Использовать ли песочницу SandboxedEnvironment.
        :param update_fields: Включить ли автоматический пересчёт полей при открытии документа в Word.
        """

        self.render(
            context,
            jinja_env=jinja_env,
            autoescape=autoescape,
            sandbox=sandbox,
            update_fields=update_fields,
        )

    def fix_tables(self, xml: str) -> etree._Element:
        """
        Приводит количество и ширину колонок <w:gridCol> в таблицах в соответствие с фактическим числом ячеек <w:tc>.

        :param xml: Отрендеренная XML-строка тела документа.

        :return: Скорректированное XML-дерево элемента <w:body>.
        """

        parser = etree.XMLParser(recover=True)
        tree = etree.fromstring(xml, parser=parser)
        ns = "{" + tree.nsmap["w"] + "}"

        for t in tree.iter(ns + "tbl"):
            tbl_grid = t.find(ns + "tblGrid")
            if tbl_grid is None:
                continue
            columns = tbl_grid.findall(ns + "gridCol")
            to_add = 0

            rows = t.findall(ns + "tr")
            for r in rows:
                cells = r.findall(ns + "tc")
                if (len(columns) + to_add) < len(cells):
                    to_add = len(cells) - len(columns)

            if to_add > 0:
                width = 0.0
                new_average = None
                for c in columns:
                    w_val = c.get(ns + "w")
                    if w_val is not None:
                        try:
                            width += float(w_val)
                        except ValueError:
                            pass
                if width > 0 and len(columns) > 0:
                    old_average = width / len(columns)
                    new_average = width / (len(columns) + to_add)
                    for c in columns:
                        w_val = c.get(ns + "w")
                        if w_val is not None:
                            c.set(
                                ns + "w",
                                str(int(float(w_val) * new_average / old_average)),
                            )
                    for _ in range(to_add):
                        etree.SubElement(
                            tbl_grid, ns + "gridCol", {ns + "w": str(int(new_average))}
                        )

            columns = tbl_grid.findall(ns + "gridCol")
            columns_len = len(columns)

            cells_len_max = 0

            def get_cell_len(total: int, cell: etree._Element) -> int:
                tc_pr = cell.find(ns + "tcPr")
                grid_span = None if tc_pr is None else tc_pr.find(ns + "gridSpan")

                if grid_span is not None:
                    val_str = grid_span.get(ns + "val")
                    if val_str and val_str.isdigit():
                        return total + int(val_str)

                return total + 1

            for r in rows:
                cells = r.findall(ns + "tc")
                cells_len = functools.reduce(get_cell_len, cells, 0)
                cells_len_max = max(cells_len_max, cells_len)

            to_remove = columns_len - cells_len_max

            if to_remove > 0 and cells_len_max > 0:
                removed_width = 0.0

                for c in columns[-to_remove:]:
                    w_val = c.get(ns + "w")
                    if w_val is not None:
                        try:
                            removed_width += float(w_val)
                        except ValueError:
                            pass
                    tbl_grid.remove(c)

                columns_left = tbl_grid.findall(ns + "gridCol")

                extra_space = 0
                if len(columns_left) > 0:
                    extra_space = int(removed_width / len(columns_left))

                for c in columns_left:
                    w_val = c.get(ns + "w")
                    if w_val is not None:
                        c.set(ns + "w", str(int(float(w_val) + extra_space)))

        return tree

    def fix_docpr_ids(self, tree: etree._Element) -> None:
        """
        Перенумеровывает идентификаторы объектов (wp:docPr, pic:cNvPr, w:object / o:OLEObject) во избежание коллизий (Issue #624).

        :param tree: XML-дерево тела документа.
        """

        nsmap = dict(docx.oxml.ns.nsmap)
        nsmap.setdefault("v", "urn:schemas-microsoft-com:vml")
        nsmap.setdefault("o", "urn:schemas-microsoft-com:office:office")
        nsmap.setdefault("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")

        for elt in tree.xpath("//wp:docPr", namespaces=nsmap):
            self.docx_ids_index += 1
            elt.attrib["id"] = str(self.docx_ids_index)

        for elt in tree.xpath("//pic:cNvPr", namespaces=nsmap):
            self.docx_ids_index += 1
            elt.attrib["id"] = str(self.docx_ids_index)

        seen_shape_ids: set[str] = set()
        seen_object_ids: set[str] = set()
        seen_ole_rids: set[str] = set()
        r_id_attr = f"{{{nsmap['r']}}}id"

        for w_obj in tree.xpath("//w:object", namespaces=nsmap):
            shapes = w_obj.xpath(".//v:shape", namespaces=nsmap)
            ole_objs = w_obj.xpath(".//o:OLEObject", namespaces=nsmap)
            if not ole_objs:
                continue

            ole = ole_objs[0]
            shape = shapes[0] if shapes else None

            shape_id = (shape.get("id") if shape is not None else None) or ole.get("ShapeID") or ""
            if not shape_id or shape_id in seen_shape_ids:
                self.docx_ids_index += 1
                shape_id = f"_x0000_i{self.docx_ids_index}"
                if shape is not None:
                    shape.set("id", shape_id)
                ole.set("ShapeID", shape_id)
            seen_shape_ids.add(shape_id)

            obj_id = ole.get("ObjectID") or ""
            if not obj_id or obj_id in seen_object_ids:
                self.docx_ids_index += 1
                obj_id = f"_{1000000000 + self.docx_ids_index}"
                ole.set("ObjectID", obj_id)
            seen_object_ids.add(obj_id)

            ole_rid = ole.get(r_id_attr)
            if ole_rid and self.docx is not None and ole_rid in self.docx._part.rels:
                if ole_rid in seen_ole_rids:
                    orig_rel = self.docx._part.rels[ole_rid]
                    if not orig_rel.is_external and orig_rel.target_part is not None:
                        orig_part = orig_rel.target_part
                        orig_name = str(orig_part.partname)
                        base, ext = os.path.splitext(orig_name)
                        self.docx_ids_index += 1
                        new_partname = PackURI(f"{base}_{self.docx_ids_index}{ext}")
                        cloned_part = Part(
                            new_partname,
                            orig_part.content_type,
                            orig_part.blob,
                            orig_part.package,
                        )
                        new_rid = self.docx._part.relate_to(cloned_part, orig_rel.reltype)
                        ole.set(r_id_attr, new_rid)
                else:
                    seen_ole_rids.add(ole_rid)

    def new_subdoc(
        self,
        docpath: str | PathLike[str] | IO[bytes] | DocxDocument | None = None,
        remove_trailing_paragraph: bool = False,
    ) -> Subdoc:
        """
        Создаёт новый поддокумент (пустой или на основе существующего файла/объекта .docx).

        :param docpath: Путь к файлу .docx, поток BytesIO, объект Document в памяти или None.
        :param remove_trailing_paragraph: Удалять ли пустой завершающий абзац при вставке.

        :return: Экземпляр поддокумента Subdoc.
        """

        from .subdoc import Subdoc

        self.init_docx(reload=False)
        return Subdoc(
            self,
            docpath=docpath,
            remove_trailing_paragraph=remove_trailing_paragraph,
        )

    @staticmethod
    def get_file_crc(file_obj: str | PathLike[str] | IO[bytes]) -> int:
        """
        Вычисляет контрольную сумму CRC32 для файла или файлоподобного объекта.

        :param file_obj: Путь к файлу или бинарный поток с методом read().

        :return: Беззнаковое 32-битное целое число CRC32.
        """

        if hasattr(file_obj, "read"):
            if hasattr(file_obj, "seek"):
                file_obj.seek(0)  # type: ignore[union-attr]
            buf = file_obj.read()  # type: ignore[union-attr]
        else:
            with open(file_obj, "rb") as fh:
                buf = fh.read()

        crc = binascii.crc32(buf) & 0xFFFFFFFF
        return crc

    @staticmethod
    def _read_bytes(file_or_path: str | PathLike[str] | IO[bytes]) -> bytes:
        """
        Считывает байты из пути к файлу или бинарного потока (с перемоткой в начало при наличии seek).

        :param file_or_path: Путь к файлу или файлоподобный объект.

        :return: Считанные байты.
        """

        if hasattr(file_or_path, "read"):
            if hasattr(file_or_path, "seek"):
                try:
                    file_or_path.seek(0)  # type: ignore[union-attr]
                except Exception:
                    pass
            return file_or_path.read()  # type: ignore[union-attr]
        with open(file_or_path, "rb") as fh:
            return fh.read()

    def replace_media(
        self,
        src_file: str | PathLike[str] | IO[bytes],
        dst_file: str | PathLike[str] | IO[bytes],
    ) -> None:
        """
        Заменяет медиафайл в документе (включая колонтитулы) другим медиафайлом по совпадению CRC32.

        :param src_file: Путь или поток исходного медиафайла-заглушки, вставленного в шаблон.
        :param dst_file: Путь или поток нового медиафайла для замены.
        """

        crc = self.get_file_crc(src_file)
        self.crc_to_new_media[crc] = self._read_bytes(dst_file)

    def replace_pic(
        self,
        embedded_file: str,
        dst_file: str | PathLike[str] | IO[bytes],
    ) -> None:
        """
        Заменяет встроенное изображение по его имени, заголовку (title) или описанию (alt text / descr).

        :param embedded_file: Базовое имя файла, заголовок или описание заменяемой картинки в шаблоне.
        :param dst_file: Путь к новому файлу изображения или бинарный поток.
        """

        self.pics_to_replace[embedded_file] = self._read_bytes(dst_file)

    def replace_embedded(
        self,
        src_file: str | PathLike[str] | IO[bytes],
        dst_file: str | PathLike[str] | IO[bytes],
    ) -> None:
        """
        Заменяет встроенный в документ объект (например, вложенный .docx) по контрольной сумме CRC32.

        :param src_file: Путь или поток исходного встроенного файла-заглушки.
        :param dst_file: Путь или поток нового встроенного файла.
        """

        crc = self.get_file_crc(src_file)
        self.crc_to_new_embedded[crc] = self._read_bytes(dst_file)

    def replace_zipname(
        self,
        zipname: str,
        dst_file: str | PathLike[str] | IO[bytes],
    ) -> None:
        """
        Заменяет внутренний файл архива .docx по его полному пути внутри ZIP-пакета.

        :param zipname: Внутренний путь к файлу в ZIP-архиве .docx.
        :param dst_file: Путь к новому файлу или бинарный поток.
        """

        self.zipname_to_replace[zipname] = self._read_bytes(dst_file)

    def reset_replacements(self) -> None:
        """
        Сбрасывает все словари запланированных замен медиафайлов, картинок и встроенных объектов.
        """

        self.crc_to_new_media: dict[int, bytes] = {}
        self.crc_to_new_embedded: dict[int, bytes] = {}
        self.zipname_to_replace: dict[str, bytes] = {}
        self.pics_to_replace: dict[str, bytes] = {}

    def post_processing(self, docx_file: IO[bytes] | str | PathLike[str]) -> None:
        """
        Выполняет постобработку сохранённого ZIP-архива .docx для замены медиафайлов и встроенных объектов.

        :param docx_file: Путь к сохранённому файлу .docx или бинарный поток.
        """

        if self.crc_to_new_media or self.crc_to_new_embedded or self.zipname_to_replace:
            if hasattr(docx_file, "read"):
                tmp_file: Any = io.BytesIO()
                DocxTemplate(docx_file).save(tmp_file)
                tmp_file.seek(0)
                docx_file.seek(0)  # type: ignore[union-attr]
                docx_file.truncate()  # type: ignore[union-attr]
                docx_file.seek(0)  # type: ignore[union-attr]
            else:
                tmp_file = f"{os.fspath(docx_file)}_docxtpl_before_replace_medias"
                if os.path.exists(tmp_file):
                    os.remove(tmp_file)
                os.rename(docx_file, tmp_file)

            with zipfile.ZipFile(tmp_file, "r") as zin:
                with zipfile.ZipFile(docx_file, "w", compression=zipfile.ZIP_DEFLATED) as zout:
                    for item in zin.infolist():
                        buf = zin.read(item.filename)
                        if item.filename in self.zipname_to_replace:
                            zout.writestr(item.filename, self.zipname_to_replace[item.filename])
                        elif (
                            item.filename.startswith("word/media/")
                            and item.CRC in self.crc_to_new_media
                        ):
                            zout.writestr(item.filename, self.crc_to_new_media[item.CRC])
                        elif (
                            item.filename.startswith("word/embeddings/")
                            and item.CRC in self.crc_to_new_embedded
                        ):
                            zout.writestr(item.filename, self.crc_to_new_embedded[item.CRC])
                        else:
                            zout.writestr(item, buf)

            if not hasattr(tmp_file, "read"):
                os.remove(tmp_file)
            if hasattr(docx_file, "read"):
                docx_file.seek(0)  # type: ignore[union-attr]

    def pre_processing(self) -> None:
        """
        Выполняет предобработку документа перед сохранением (замена изображений через replace_pic).
        """

        if self.pics_to_replace:
            self._replace_pics()

    def _replace_pics(self) -> None:
        """
        Заменяет изображения в основной части документа и всех колонтитулах согласно словарю pics_to_replace.
        """

        assert self.docx is not None
        replaced_pics = {key: False for key in self.pics_to_replace}

        part = self.docx.part
        self._replace_docx_part_pics(part, replaced_pics)

        for relid, rel in part.rels.items():
            if rel.reltype in (REL_TYPE.HEADER, REL_TYPE.FOOTER):
                self._replace_docx_part_pics(rel.target_part, replaced_pics)

        if not self.allow_missing_pics:
            for img_id, replaced in replaced_pics.items():
                if not replaced:
                    raise ValueError(
                        f"Изображение {img_id} не найдено в шаблоне docx (Picture {img_id} not found in the docx template)"
                    )

    def get_pic_map(self) -> dict[str, Any]:
        """
        Возвращает словарь сопоставления найденных при замене изображений их частям в пакете.

        :return: Словарь {имя_файла: (target_ref, target_part)}.
        """

        return self.pic_map

    def _replace_docx_part_pics(
        self,
        doc_part: Any,
        replaced_pics: dict[str, bool],
    ) -> None:
        """
        Ищет и заменяет изображения в конкретной части документа по атрибутам pic:cNvPr, wp:docPr и имени медиафайла (Issue #618).

        :param doc_part: Часть документа (основная или колонтитул).
        :param replaced_pics: Словарь флагов успешной замены для каждого искомого идентификатора изображения.
        """

        et = etree.fromstring(doc_part.blob)

        part_map: dict[str, Any] = {}

        gds = et.xpath("//a:graphic/a:graphicData", namespaces=docx.oxml.ns.nsmap)
        for gd in gds:
            rel = None
            try:
                if gd.attrib.get("uri") == docx.oxml.ns.nsmap["pic"]:
                    blip_list = gd.xpath(
                        "pic:pic/pic:blipFill/a:blip", namespaces=docx.oxml.ns.nsmap
                    )
                    if not blip_list:
                        continue
                    blip = blip_list[0]
                    dest = blip.xpath("@r:embed", namespaces=docx.oxml.ns.nsmap)
                    if len(dest) > 0:
                        rel = dest[0]
                    else:
                        continue
                else:
                    continue

                non_visual_properties = "pic:pic/pic:nvPicPr/pic:cNvPr/"
                filenames = gd.xpath(
                    f"{non_visual_properties}@name", namespaces=docx.oxml.ns.nsmap
                )
                filename = filenames[0] if filenames else ""

                titles = gd.xpath(
                    f"{non_visual_properties}@title", namespaces=docx.oxml.ns.nsmap
                )
                title = titles[0] if titles else ""

                descriptions = gd.xpath(
                    f"{non_visual_properties}@descr", namespaces=docx.oxml.ns.nsmap
                )
                description = descriptions[0] if descriptions else ""

                doc_pr_list = gd.xpath(
                    "ancestor::wp:inline[1]/wp:docPr | ancestor::wp:anchor[1]/wp:docPr",
                    namespaces=docx.oxml.ns.nsmap,
                )
                doc_pr_name = doc_pr_list[0].get("name", "") if doc_pr_list else ""
                doc_pr_title = doc_pr_list[0].get("title", "") if doc_pr_list else ""
                doc_pr_descr = doc_pr_list[0].get("descr", "") if doc_pr_list else ""

                target_ref = doc_part.rels[rel].target_ref
                target_part = doc_part.rels[rel].target_part
                target_basename = os.path.basename(target_ref)

                map_key = filename or doc_pr_name or target_basename
                part_map[map_key] = (target_ref, target_part)

                candidate_ids = {
                    val
                    for val in (
                        filename,
                        title,
                        description,
                        doc_pr_name,
                        doc_pr_title,
                        doc_pr_descr,
                        target_basename,
                        os.path.basename(filename) if filename else "",
                        os.path.basename(description) if description else "",
                        os.path.basename(doc_pr_descr) if doc_pr_descr else "",
                    )
                    if val
                }
                candidate_ids |= {urllib.parse.unquote(val) for val in list(candidate_ids)}

                for img_id, img_data in self.pics_to_replace.items():
                    if img_id in candidate_ids or urllib.parse.unquote(img_id) in candidate_ids:
                        target_part._blob = img_data
                        replaced_pics[img_id] = True

            except Exception:
                continue

        self.pic_map.update(part_map)

    def build_url_id(self, url: str) -> str:
        """
        Создаёт внешнюю связь гиперссылки в текущем документе и возвращает её идентификатор rId (для использования в RichText).

        :param url: Целевой URL-адрес гиперссылки.

        :return: Строковый идентификатор связи (например, 'rId9').
        """

        self.init_docx(reload=False)
        assert self.docx is not None
        return self.docx._part.relate_to(url, REL_TYPE.HYPERLINK, is_external=True)

    def save(
        self,
        filename: IO[bytes] | str | PathLike[str],
        *args: Any,
        **kwargs: Any,
    ) -> None:
        """
        Сохраняет сгенерированный документ .docx по указанному пути или в бинарный поток.

        :param filename: Путь к выходному файлу .docx или файлоподобный объект.
        :param args: Дополнительные позиционные аргументы для python-docx save().
        :param kwargs: Дополнительные именованные аргументы для python-docx save().
        """

        if not self.is_saved and not self.is_rendered:
            if self.docx is None:
                self.init_docx()
        assert self.docx is not None

        if self._custom_properties is not None:
            self._custom_properties.sync()
            self.update_docproperty_fields()

        self.pre_processing()
        self.docx.save(filename, *args, **kwargs)
        self.post_processing(filename)
        self.is_saved = True

    async def save_async(
        self,
        filename: IO[bytes] | str | PathLike[str],
        *args: Any,
        **kwargs: Any,
    ) -> None:
        """
        Асинхронно сохраняет сгенерированный документ .docx в корутинах asyncio (Issue #611).

        :param filename: Путь к выходному файлу .docx или файлоподобный объект.
        :param args: Дополнительные позиционные аргументы.
        :param kwargs: Дополнительные именованные аргументы.
        """

        self.save(filename, *args, **kwargs)

    def get_undeclared_template_variables(
        self,
        jinja_env: Environment | None = None,
        context: dict[str, Any] | None = None,
    ) -> set[str]:
        """
        Возвращает множество всех необъявленных переменных шаблона Jinja2 в документе (включая тело, колонтитулы, сноски и свойства).

        :param jinja_env: Пользовательское окружение Jinja2 или None.
        :param context: Опциональный словарь уже переданных переменных контекста.

        :return: Множество строковых имён переменных.
        """

        if hasattr(self.template_file, "seek"):
            try:
                self.template_file.seek(0)  # type: ignore[union-attr]
            except Exception:
                pass
        temp_doc = Document(self.template_file)

        xml = self.xml_to_string(temp_doc._element.body)
        xml = self.patch_xml(xml)

        for uri in [self.HEADER_URI, self.FOOTER_URI]:
            for rel_key, val in temp_doc._part.rels.items():
                if (val.reltype == uri) and (val.target_part.blob):
                    _xml = self.xml_to_string(parse_xml(val.target_part.blob))
                    xml += self.patch_xml(_xml)

        if temp_doc.part.package is not None:
            target_types = {
                self.FOOTNOTES_CONTENT_TYPE,
                self.ENDNOTES_CONTENT_TYPE,
                self.COMMENTS_CONTENT_TYPE,
            }
            for part in temp_doc.part.package.parts:
                if part.content_type in target_types and part.blob:
                    raw_xml = (
                        part.blob.decode("utf-8")
                        if isinstance(part.blob, bytes)
                        else part.blob
                    )
                    xml += self.patch_xml(raw_xml)

        for prop in CORE_STRING_PROPERTIES:
            val_str = getattr(temp_doc.core_properties, prop, "")
            if val_str and ("{{" in val_str or "{%" in val_str):
                xml += val_str
            if self.docx is not None and not self.is_rendered:
                cur_val = getattr(self.docx.core_properties, prop, "")
                if cur_val and cur_val != val_str and ("{{" in cur_val or "{%" in cur_val):
                    xml += cur_val

        for val_str in get_app_properties_dict(temp_doc).values():
            if val_str and ("{{" in val_str or "{%" in val_str):
                xml += val_str
        if self.docx is not None and not self.is_rendered:
            for val_str in get_app_properties_dict(self.docx).values():
                if val_str and ("{{" in val_str or "{%" in val_str):
                    xml += val_str

        for val_any in CustomProperties(temp_doc).values():
            if isinstance(val_any, str) and ("{{" in val_any or "{%" in val_any):
                xml += val_any
        if self._custom_properties is not None and not self.is_rendered:
            for val_any in self._custom_properties.values():
                if isinstance(val_any, str) and ("{{" in val_any or "{%" in val_any):
                    xml += val_any

        env = self._get_or_create_jinja_env(jinja_env=jinja_env)

        parse_content = env.parse(xml)
        all_variables = meta.find_undeclared_variables(parse_content)

        if context is not None:
            provided_variables = set(context.keys())
            return all_variables - provided_variables

        return all_variables

    def _collect_template_parts_for_inspection(self) -> list[tuple[str, str, bool]]:
        """
        Собирает все части шаблона для синтаксической проверки и поиска переменных.

        :return: Список кортежей (имя_части, строка_содержимого, является_ли_xml).
        """

        if hasattr(self.template_file, "seek"):
            try:
                self.template_file.seek(0)  # type: ignore[union-attr]
            except Exception:
                pass
        temp_doc = Document(self.template_file)
        parts: list[tuple[str, str, bool]] = []

        # 1. Основное тело документа
        body_xml = self.xml_to_string(temp_doc._element.body)
        parts.append(("body", self.patch_xml(body_xml), True))

        # 2. Верхние и нижние колонтитулы
        for label, uri in (("header", self.HEADER_URI), ("footer", self.FOOTER_URI)):
            for rel_key, val in temp_doc._part.rels.items():
                if (val.reltype == uri) and (val.target_part.blob):
                    part_xml = self.xml_to_string(parse_xml(val.target_part.blob))
                    parts.append((f"{label} ({rel_key})", self.patch_xml(part_xml), True))

        # 3. Сноски, концевые сноски и примечания
        if temp_doc.part.package is not None:
            type_labels = {
                self.FOOTNOTES_CONTENT_TYPE: "footnotes",
                self.ENDNOTES_CONTENT_TYPE: "endnotes",
                self.COMMENTS_CONTENT_TYPE: "comments",
            }
            seen_partnames: set[str] = set()
            for part in temp_doc.part.package.parts:
                partname = str(getattr(part, "partname", id(part)))
                if partname in seen_partnames:
                    continue
                if part.content_type in type_labels and part.blob:
                    seen_partnames.add(partname)
                    raw_xml = (
                        part.blob.decode("utf-8")
                        if isinstance(part.blob, bytes)
                        else part.blob
                    )
                    parts.append((type_labels[part.content_type], self.patch_xml(raw_xml), True))

        # 4. Стандартные свойства Core (из шаблона + установленные до рендеринга)
        core_sources = [temp_doc]
        if self.docx is not None and not self.is_rendered:
            core_sources.append(self.docx)
        seen_core: set[tuple[str, str]] = set()
        for src_doc in core_sources:
            for prop in CORE_STRING_PROPERTIES:
                val_str = getattr(src_doc.core_properties, prop, "")
                if val_str and ("{{" in val_str or "{%" in val_str) and (prop, val_str) not in seen_core:
                    seen_core.add((prop, val_str))
                    parts.append((f"core_properties.{prop}", val_str, False))

        # 5. Расширенные свойства App
        seen_app: set[tuple[str, str]] = set()
        for src_doc in core_sources:
            for key, val_str in get_app_properties_dict(src_doc).items():
                if val_str and ("{{" in val_str or "{%" in val_str) and (key, val_str) not in seen_app:
                    seen_app.add((key, val_str))
                    parts.append((f"app_properties.{key}", val_str, False))

        # 6. Пользовательские свойства Custom
        seen_custom: set[tuple[str, str]] = set()
        custom_dicts = [CustomProperties(temp_doc).to_dict()]
        if self._custom_properties is not None and not self.is_rendered:
            custom_dicts.append(self._custom_properties.to_dict())
        for cdict in custom_dicts:
            for key, val_any in cdict.items():
                if isinstance(val_any, str) and ("{{" in val_any or "{%" in val_any) and (key, val_any) not in seen_custom:
                    seen_custom.add((key, val_any))
                    parts.append((f"custom_properties.{key}", val_any, False))

        return parts

    @staticmethod
    def _extract_error_context_lines(src_xml: str, lineno: int | None, is_xml: bool) -> list[str]:
        """
        Извлекает читаемые текстовые строки абзацев вокруг строки с синтаксической ошибкой Jinja2.

        :param src_xml: Подготовленная строка шаблона (XML или текст свойства).
        :param lineno: Номер строки ошибки из исключения TemplateError.
        :param is_xml: Признак того, что строка является XML-документом Word.

        :return: Список очищенных от XML-тегов строк контекста.
        """

        if not is_xml:
            return [src_xml.strip()] if src_xml.strip() else []
        if lineno is None:
            return []
        lines = src_xml.splitlines()
        line_idx = max(0, lineno - 1)
        window = lines[max(0, line_idx - 2) : min(len(lines), line_idx + 3)]
        return [
            re.sub(r"<[^>]+>", "", line).strip()
            for line in window
            if re.sub(r"<[^>]+>", "", line).strip()
        ]

    def validate_template(
        self,
        jinja_env: Environment | None = None,
        raise_on_error: bool = False,
    ) -> list[TemplateValidationError]:
        """
        Проверяет синтаксис всех частей шаблона (тело документа, верхние и нижние колонтитулы,
        сноски, примечания и свойства документа) без остановки на первой ошибке.

        :param jinja_env: Пользовательское окружение Jinja2 или None.
        :param raise_on_error: Если True — выбрасывает исключение TemplateError при наличии ошибок.

        :return: Список объектов TemplateValidationError (пустой список, если шаблон корректен).
        """

        env = self._get_or_create_jinja_env(jinja_env=jinja_env)
        errors: list[TemplateValidationError] = []

        for part_name, content, is_xml in self._collect_template_parts_for_inspection():
            prepared = re.sub(r"<w:p([ >])", r"\n<w:p\1", content) if is_xml else content
            try:
                env.parse(prepared)
            except TemplateError as err:
                lineno = getattr(err, "lineno", None)
                ctx_lines = self._extract_error_context_lines(prepared, lineno, is_xml)
                errors.append(
                    TemplateValidationError(
                        part_name=part_name,
                        message=str(err.message if hasattr(err, "message") and err.message else err),
                        lineno=lineno,
                        context_lines=ctx_lines,
                    )
                )

        if raise_on_error and errors:
            combined_msg = "\n\n".join(str(e) for e in errors)
            raise TemplateError(combined_msg)

        return errors

    def is_template_valid(self, jinja_env: Environment | None = None) -> bool:
        """
        Быстро проверяет, не содержит ли шаблон синтаксических ошибок Jinja2.

        :param jinja_env: Пользовательское окружение Jinja2 или None.

        :return: True, если все части шаблона синтаксически корректны, иначе False.
        """

        return len(self.validate_template(jinja_env=jinja_env)) == 0

    def check_context(
        self,
        context: dict[str, Any],
        jinja_env: Environment | None = None,
    ) -> ContextCheckResult:
        """
        Проверяет полноту словаря контекста и синтаксическую корректность шаблона перед рендерингом.
        Возвращает отчёт со списками недостающих переменных (missing), неиспользуемых ключей (unused),
        всех требуемых переменных (required) и синтаксических ошибок (syntax_errors).

        :param context: Словарь переменных контекста, который планируется передать в render().
        :param jinja_env: Пользовательское окружение Jinja2 или None.

        :return: Экземпляр ContextCheckResult с детальной информацией о проверке.
        """

        env = self._get_or_create_jinja_env(jinja_env=jinja_env)
        required_vars: set[str] = set()
        syntax_errors: list[TemplateValidationError] = []

        for part_name, content, is_xml in self._collect_template_parts_for_inspection():
            prepared = re.sub(r"<w:p([ >])", r"\n<w:p\1", content) if is_xml else content
            try:
                ast = env.parse(prepared)
                required_vars |= meta.find_undeclared_variables(ast)
            except TemplateError as err:
                lineno = getattr(err, "lineno", None)
                ctx_lines = self._extract_error_context_lines(prepared, lineno, is_xml)
                syntax_errors.append(
                    TemplateValidationError(
                        part_name=part_name,
                        message=str(err.message if hasattr(err, "message") and err.message else err),
                        lineno=lineno,
                        context_lines=ctx_lines,
                    )
                )

        return ContextCheckResult(
            required=required_vars,
            provided=set(context.keys()),
            syntax_errors=syntax_errors,
        )

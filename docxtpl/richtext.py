# -*- coding: utf-8 -*-
"""
Модуль для генерации форматированного текста (RichText) и абзацев (RichTextParagraph) внутри шаблонов Word.
"""
from __future__ import annotations

from html import escape
from typing import Any


class RichText:
    """ Класс для динамической генерации форматированного текста (прогонов <w:r>) внутри абзаца. """

    def __init__(self, text: Any = None, **text_prop: Any) -> None:
        """
        Инициализация объекта форматированного текста RichText.

        :param text: Начальный текст или другой объект RichText (если None — создаётся пустой объект).
        :param text_prop: Именованные параметры форматирования текста (передаются в метод add).
        """

        self.xml: str = ""
        if text is not None:
            self.add(text, **text_prop)

    def add(
        self,
        text: Any,
        style: str | None = None,
        color: str | None = None,
        highlight: str | None = None,
        size: int | str | None = None,
        subscript: bool | None = None,
        superscript: bool | None = None,
        bold: bool = False,
        italic: bool = False,
        underline: bool | str = False,
        strike: bool = False,
        font: str | None = None,
        url_id: str | None = None,
        rtl: bool = False,
        lang: str | None = None,
    ) -> None:
        """
        Добавляет фрагмент текста с заданными свойствами форматирования во внутренний XML-буфер.

        :param text: Добавляемый текст, число или другой экземпляр RichText.
        :param style: Имя символьного стиля Word (w:rStyle).
        :param color: Цвет текста в формате HEX (например, 'ff0000' или '#ff0000').
        :param highlight: Цвет выделения (фона) текста в формате HEX.
        :param size: Размер шрифта в полупунктах (дважды размер в пт) или строка.
        :param subscript: Флаг нижнего индекса.
        :param superscript: Флаг верхнего индекса.
        :param bold: Флаг полужирного начертания.
        :param italic: Флаг курсивного начертания.
        :param underline: Флаг или тип подчёркивания ('single', 'double', 'thick', 'dotted', 'dash', 'dotDash', 'dotDotDash', 'wave').
        :param strike: Флаг зачёркнутого текста.
        :param font: Имя шрифта (опционально с префиксом региона, например 'eastAsia:微软雅黑').
        :param url_id: Идентификатор связи гиперссылки (полученный через tpl.build_url_id(url)).
        :param rtl: Флаг направления письма справа налево (Right-to-Left).
        :param lang: Код языка проверки правописания (например, 'ru-RU', 'en-US').
        """

        # Если передан другой объект RichText, просто конкатенируем его XML
        if isinstance(text, RichText):
            self.xml += text.xml
            return

        # Приводим нестроковые типы (int, float, dict и т.д.) к строке
        if not isinstance(text, (str, bytes)):
            text = str(text)
        if not isinstance(text, str):
            text = text.decode("utf-8", errors="ignore")
        text = escape(text)

        prop = ""

        if style:
            prop += f'<w:rStyle w:val="{style}"/>'
        if font:
            regional_font = ""
            if ":" in font:
                region, font = font.split(":", 1)
                regional_font = f' w:{region}="{font}"'
            prop += f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}"{regional_font}/>'
        if color:
            if color[0] == "#":
                color = color[1:]
            prop += f'<w:color w:val="{color}"/>'
        if highlight:
            if highlight[0] == "#":
                highlight = highlight[1:]
            prop += f'<w:shd w:fill="{highlight}"/>'
        if size:
            prop += f'<w:sz w:val="{size}"/>'
            prop += f'<w:szCs w:val="{size}"/>'
        if subscript:
            prop += '<w:vertAlign w:val="subscript"/>'
        if superscript:
            prop += '<w:vertAlign w:val="superscript"/>'
        if bold:
            prop += "<w:b/>"
            if rtl:
                prop += "<w:bCs/>"
        if italic:
            prop += "<w:i/>"
            if rtl:
                prop += "<w:iCs/>"
        if underline:
            if underline not in [
                "single",
                "double",
                "thick",
                "dotted",
                "dash",
                "dotDash",
                "dotDotDash",
                "wave",
            ]:
                underline = "single"
            prop += f'<w:u w:val="{underline}"/>'
        if strike:
            prop += "<w:strike/>"
        if rtl:
            prop += '<w:rtl w:val="true"/>'
        if lang:
            prop += f'<w:lang w:val="{lang}"/>'

        xml = "<w:r>"
        if prop:
            xml += f"<w:rPr>{prop}</w:rPr>"
        xml += f'<w:t xml:space="preserve">{text}</w:t></w:r>'
        if url_id:
            xml = f'<w:hyperlink r:id="{url_id}" w:tgtFrame="_blank">{xml}</w:hyperlink>'
        self.xml += xml

    def __bool__(self) -> bool:
        """
        Проверяет, содержит ли объект RichText какую-либо XML-разметку (Issue #601).

        :return: True, если объект не пустой, иначе False.
        """

        return bool(self.xml)

    def __eq__(self, other: object) -> bool:
        """
        Сравнивает два объекта RichText по их внутренней XML-строке.

        :param other: Другой объект для сравнения.

        :return: True, если XML-разметка совпадает.
        """

        if isinstance(other, RichText):
            return self.xml == other.xml
        return False

    def __unicode__(self) -> str:
        """
        Возвращает XML-представление форматированного текста (для обратной совместимости).

        :return: Строка XML-разметки.
        """

        return self.xml

    def __str__(self) -> str:
        """
        Возвращает строковое XML-представление форматированного текста.

        :return: Строка XML-разметки.
        """

        return self.xml

    def __html__(self) -> str:
        """
        Возвращает XML-представление как безопасную разметку для Jinja2 autoescape.

        :return: Строка XML-разметки.
        """

        return self.xml


class RichTextParagraph:
    """ Класс для динамической генерации форматированных абзацев (<w:p>) с поддержкой стилей абзаца. """

    def __init__(self, text: Any = None, **text_prop: Any) -> None:
        """
        Инициализация объекта форматированных абзацев RichTextParagraph.

        :param text: Начальный текст или объект RichText (если None — создаётся пустой объект).
        :param text_prop: Дополнительные свойства абзаца (например, parastyle).
        """

        self.xml: str = ""
        if text is not None:
            self.add(text, **text_prop)

    def add(
        self,
        text: Any,
        parastyle: str | None = None,
    ) -> None:
        """
        Добавляет новый абзац <w:p> с указанным текстом и стилем абзаца.

        :param text: Текст или экземпляр RichText для вставки в абзац.
        :param parastyle: Имя стиля абзаца Word (w:pStyle).
        """

        # Если передан не RichText, оборачиваем в RichText
        if not isinstance(text, RichText):
            text = RichText(text)

        prop = ""
        if parastyle:
            prop += f'<w:pStyle w:val="{parastyle}"/>'

        xml = "<w:p>"
        if prop:
            xml += f"<w:pPr>{prop}</w:pPr>"
        xml += text.xml
        xml += "</w:p>"
        self.xml += xml

    def __bool__(self) -> bool:
        """
        Проверяет, содержит ли объект RichTextParagraph добавленные абзацы (Issue #601).

        :return: True, если объект не пустой, иначе False.
        """

        return bool(self.xml)

    def __eq__(self, other: object) -> bool:
        """
        Сравнивает два объекта RichTextParagraph по их XML-разметке.

        :param other: Другой объект для сравнения.

        :return: True, если XML-разметка совпадает.
        """

        if isinstance(other, RichTextParagraph):
            return self.xml == other.xml
        return False

    def __unicode__(self) -> str:
        """
        Возвращает XML-представление абзацев (для обратной совместимости).

        :return: Строка XML-разметки.
        """

        return self.xml

    def __str__(self) -> str:
        """
        Возвращает строковое XML-представление абзацев.

        :return: Строка XML-разметки.
        """

        return self.xml

    def __html__(self) -> str:
        """
        Возвращает XML-представление как безопасную разметку для Jinja2 autoescape.

        :return: Строка XML-разметки.
        """

        return self.xml


R = RichText
RP = RichTextParagraph

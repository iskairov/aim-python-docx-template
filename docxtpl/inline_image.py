# -*- coding: utf-8 -*-
"""
Модуль для генерации встраиваемых изображений (InlineImage) внутри шаблонов Word.
"""
from __future__ import annotations

from os import PathLike
from typing import TYPE_CHECKING, Any, IO

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

if TYPE_CHECKING:
    from docx.shared import Length
    from .template import DocxTemplate


class InlineImage:
    """ Класс для быстрой генерации встраиваемых в текст изображений без использования Subdoc. """

    tpl: DocxTemplate | None = None
    image_descriptor: str | PathLike[str] | IO[bytes] | None = None
    width: Length | int | None = None
    height: Length | int | None = None
    anchor: str | None = None
    title: str | None = None
    descr: str | None = None

    def __init__(
        self,
        tpl: DocxTemplate,
        image_descriptor: str | PathLike[str] | IO[bytes],
        width: Length | int | None = None,
        height: Length | int | None = None,
        anchor: str | None = None,
        title: str | None = None,
        descr: str | None = None,
    ) -> None:
        """
        Инициализация встраиваемого изображения.

        :param tpl: Экземпляр текущего шаблона DocxTemplate.
        :param image_descriptor: Путь к файлу изображения или файлоподобный поток (BytesIO).
        :param width: Ширина изображения (например, Mm(20), Inches(1), Pt(50)).
        :param height: Высота изображения (например, Mm(10), Inches(1), Pt(50)).
        :param anchor: URL-адрес для создания гиперссылки при клике на изображение.
        :param title: Заголовок изображения (атрибут title в свойствах объекта Word).
        :param descr: Описание / замещающий текст (alt text) изображения.
        """

        self.tpl, self.image_descriptor = tpl, image_descriptor
        self.width, self.height = width, height
        self.anchor = anchor
        self.title = str(title) if title is not None else None
        self.descr = str(descr) if descr is not None else None

    def _add_hyperlink(self, run: Any, url: str, part: Any) -> Any:
        """
        Добавляет гиперссылку на графический элемент изображения.

        :param run: XML-элемент встраиваемого рисунка (<wp:inline>).
        :param url: Целевой URL гиперссылки.
        :param part: Текущая рендерящаяся часть документа для регистрации связи.

        :return: Модифицированный XML-элемент рисунка с гиперссылкой.
        """

        # Создаём внешнюю связь для гиперссылки в текущей части документа
        r_id = part.relate_to(
            url,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
            is_external=True,
        )

        # Находим элементы <wp:docPr> и <pic:cNvPr> для привязки клика по ссылке
        doc_pr = run.xpath(".//wp:docPr")[0]
        cnv_pr = run.xpath(".//pic:cNvPr")[0]

        # Создаём элементы <a:hlinkClick> с идентификатором созданной связи
        hlink_click1 = OxmlElement("a:hlinkClick")
        hlink_click1.set(qn("r:id"), r_id)
        hlink_click2 = OxmlElement("a:hlinkClick")
        hlink_click2.set(qn("r:id"), r_id)

        # Вставляем элементы гиперссылки внутрь свойств объекта и рисунка
        doc_pr.append(hlink_click1)
        cnv_pr.append(hlink_click2)

        return run

    def _set_doc_properties(self, run: Any) -> None:
        """
        Устанавливает заголовок (title) и описание (descr) в XML-элементах изображения.

        :param run: XML-элемент рисунка (<wp:inline>).
        """

        doc_pr = run.xpath(".//wp:docPr")[0]
        cnv_pr = run.xpath(".//pic:cNvPr")[0]

        for elt in (doc_pr, cnv_pr):
            if self.title is not None:
                elt.set("title", self.title)
            if self.descr is not None:
                elt.set("descr", self.descr)

    def _insert_image(self) -> str:
        """
        Формирует XML-разметку для вставки изображения внутрь текущего текстового прогона.

        :return: Строка XML, закрывающая текущий <w:t>/<w:r>, вставляющая <w:drawing> и открывающая новый <w:r><w:t>.
        """

        if self.tpl is None or self.tpl.current_rendering_part is None:
            raise RuntimeError("Невозможно вставить изображение вне контекста рендеринга шаблона DocxTemplate.")

        # Если image_descriptor — поток BytesIO, сбрасываем указатель чтения в начало при повторном рендеринге
        if hasattr(self.image_descriptor, "seek"):
            try:
                self.image_descriptor.seek(0)  # type: ignore[union-attr]
            except Exception:
                pass

        pic = self.tpl.current_rendering_part.new_pic_inline(
            self.image_descriptor,
            self.width,
            self.height,
        )
        if self.title is not None or self.descr is not None:
            self._set_doc_properties(pic)
        if self.anchor:
            if pic.xpath(".//a:blip"):
                pic = self._add_hyperlink(
                    pic, self.anchor, self.tpl.current_rendering_part
                )
        pic_xml = pic.xml

        return (
            f"</w:t></w:r><w:r><w:drawing>{pic_xml}</w:drawing></w:r><w:r>"
            '<w:t xml:space="preserve">'
        )

    def __unicode__(self) -> str:
        """
        Возвращает XML-разметку изображения (для обратной совместимости).

        :return: XML-строка со встроенным изображением.
        """

        return self._insert_image()

    def __str__(self) -> str:
        """
        Возвращает строковую XML-разметку изображения.

        :return: XML-строка со встроенным изображением.
        """

        return self._insert_image()

    def __html__(self) -> str:
        """
        Возвращает XML-разметку изображения как безопасную строку при включённом autoescape в Jinja2.

        :return: XML-строка со встроенным изображением.
        """

        return self._insert_image()

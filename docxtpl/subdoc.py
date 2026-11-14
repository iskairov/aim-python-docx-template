# -*- coding: utf-8 -*-
"""
Модуль для создания и встраивания поддокументов (Subdoc) в основной документ Word.
"""
from __future__ import annotations

from copy import deepcopy
from os import PathLike
import re
from typing import TYPE_CHECKING, Any, IO

from docx import Document
from docx.document import Document as DocxDocument
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.oxml import CT_SectPr
import docx.oxml.ns
from lxml import etree

from .properties import CustomProperties

if TYPE_CHECKING:
    from .template import DocxTemplate

# Пытаемся импортировать docxcompose, если он установлен и совместим с окружением
try:
    from docxcompose.composer import Composer as _BaseComposer
    from docxcompose.utils import NS as _DOCXCOMPOSE_NS, xpath as _docxcompose_xpath

    HAS_DOCXCOMPOSE = True
except Exception:
    _BaseComposer = object  # type: ignore[assignment,misc]
    HAS_DOCXCOMPOSE = False

WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
DGM_NS = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"


class SubdocComposer(_BaseComposer):  # type: ignore[misc]
    """ Компоновщик для переноса связанных частей, стилей, списков и медиа из поддокумента в мастер-документ. """

    def __init__(self, doc: DocxDocument) -> None:
        """
        Инициализация компоновщика поддокументов.

        :param doc: Целевой (основной) документ python-docx.
        """

        self.doc = doc
        if HAS_DOCXCOMPOSE:
            try:
                super().__init__(doc)
                self._use_docxcompose = True
                return
            except Exception:
                pass
        self._use_docxcompose = False

    def attach_parts(self, doc: DocxDocument, remove_property_fields: bool = True) -> None:
        """
        Прикрепляет части поддокумента (стили, списки, изображения, диаграммы) к основному документу.

        :param doc: Поддокумент python-docx, ресурсы которого нужно перенести.
        :param remove_property_fields: Разворачивать ли поля пользовательских свойств в текст.
        """

        if self._use_docxcompose:
            self._attach_parts_docxcompose(doc, remove_property_fields=remove_property_fields)
        else:
            self._attach_parts_native(doc, remove_property_fields=remove_property_fields)

    def _attach_parts_docxcompose(
        self,
        doc: DocxDocument,
        remove_property_fields: bool = True,
    ) -> None:
        """
        Прикрепляет части поддокумента с использованием библиотеки docxcompose.

        :param doc: Поддокумент python-docx.
        :param remove_property_fields: Разворачивать ли поля пользовательских свойств.
        """

        self.reset_reference_mapping()

        # Удаляем поля кастомных свойств, сохраняя их текстовые значения
        if remove_property_fields:
            cprops = CustomProperties(doc)
            for name in cprops.keys():
                cprops.dissolve_fields(name)

        self._create_style_id_mapping(doc)

        for element in doc.element.body:
            if isinstance(element, CT_SectPr):
                continue
            self.add_referenced_parts(doc.part, self.doc.part, element)
            self.add_styles(doc, element)
            self.add_numberings(doc, element)
            self.restart_first_numbering(doc, element)
            self.add_images(doc, element)
            self.add_diagrams(doc, element)
            self.add_shapes(doc, element)
            self.add_footnotes(doc, element)
            self.remove_header_and_footer_references(doc, element)

        self.add_styles_from_other_parts(doc)
        self.renumber_bookmarks()
        self.renumber_docpr_ids()
        self.renumber_nvpicpr_ids()
        self.fix_section_types(doc)

    def _attach_parts_native(
        self,
        doc: DocxDocument,
        remove_property_fields: bool = True,
    ) -> None:
        """
        Встроенная реализация переноса всех частей поддокумента без зависимости от docxcompose.

        :param doc: Поддокумент python-docx.
        :param remove_property_fields: Разворачивать ли поля пользовательских свойств.
        """

        # 1. Разворачиваем поля пользовательских свойств, если требуется
        if remove_property_fields:
            cprops = CustomProperties(doc)
            cprops.update_fields_in_element(doc.element.body)
            cprops.dissolve_fields()

        # 2. Копируем отсутствующие стили из поддокумента в мастер-документ
        self._copy_styles_native(doc)

        # 3. Копируем определения списков и нумерации (abstractNum и num) с переназначением numId
        self._copy_numberings_native(doc)

        # 4. Переносим все внешние и внутренние связи (изображения, гиперссылки, диаграммы и пр.)
        for element in doc.element.body:
            if isinstance(element, CT_SectPr):
                continue
            self._copy_element_relationships_native(doc, element)
            self._remove_header_footer_refs_native(element)

        # 5. Перенумеровываем закладки и идентификаторы графических объектов во избежание конфликтов
        self._renumber_ids_native(doc)

    def _copy_styles_native(self, src_doc: DocxDocument) -> None:
        """
        Копирует отсутствующие стили из исходного поддокумента в основной документ.

        :param src_doc: Исходный поддокумент.
        """

        try:
            dst_styles_el = self.doc.styles.element
            src_styles_el = src_doc.styles.element
        except Exception:
            return

        w_style_id = f"{{{WORD_NS}}}styleId"
        existing_ids = {
            s.get(w_style_id)
            for s in dst_styles_el.findall(f"{{{WORD_NS}}}style")
            if s.get(w_style_id)
        }

        for style_el in src_styles_el.findall(f"{{{WORD_NS}}}style"):
            style_id = style_el.get(w_style_id)
            if style_id and style_id not in existing_ids:
                dst_styles_el.append(deepcopy(style_el))
                existing_ids.add(style_id)

    def _copy_numberings_native(self, src_doc: DocxDocument) -> None:
        """
        Копирует определения многоуровневых, маркированных и нумерованных списков в основной документ.

        :param src_doc: Исходный поддокумент.
        """

        w_ns = f"{{{WORD_NS}}}"
        num_id_els = list(src_doc.element.body.iter(f"{w_ns}numId"))
        if not num_id_els:
            return

        try:
            src_num_part = src_doc.part.numbering_part
            dst_num_part = self.doc.part.numbering_part
        except Exception:
            return

        src_num_el = src_num_part.element
        dst_num_el = dst_num_part.element

        # Находим максимальные используемые идентификаторы abstractNumId и numId в основном документе
        existing_abstract_ids = [
            int(el.get(f"{w_ns}abstractNumId", "0"))
            for el in dst_num_el.findall(f"{w_ns}abstractNum")
            if el.get(f"{w_ns}abstractNumId", "").isdigit()
        ]
        next_abstract_id = (max(existing_abstract_ids) + 1) if existing_abstract_ids else 1

        existing_num_ids = [
            int(el.get(f"{w_ns}numId", "0"))
            for el in dst_num_el.findall(f"{w_ns}num")
            if el.get(f"{w_ns}numId", "").isdigit()
        ]
        next_num_id = (max(existing_num_ids) + 1) if existing_num_ids else 1

        abstract_map: dict[str, str] = {}
        for abs_el in src_num_el.findall(f"{w_ns}abstractNum"):
            old_abs_id = abs_el.get(f"{w_ns}abstractNumId")
            if old_abs_id is None:
                continue
            new_abs_id = str(next_abstract_id)
            next_abstract_id += 1
            abstract_map[old_abs_id] = new_abs_id

            cloned_abs = deepcopy(abs_el)
            cloned_abs.set(f"{w_ns}abstractNumId", new_abs_id)

            # Удаляем nsid, чтобы Word не сливал разные списки с одинаковым nsid
            nsid_el = cloned_abs.find(f"{w_ns}nsid")
            if nsid_el is not None:
                cloned_abs.remove(nsid_el)

            # Вставляем abstractNum перед первым элементом <w:num> согласно схеме OpenXML
            first_num = dst_num_el.find(f"{w_ns}num")
            if first_num is not None:
                dst_num_el.insert(dst_num_el.index(first_num), cloned_abs)
            else:
                dst_num_el.append(cloned_abs)

        num_map: dict[str, str] = {}
        for num_el in src_num_el.findall(f"{w_ns}num"):
            old_num_id = num_el.get(f"{w_ns}numId")
            if old_num_id is None:
                continue
            new_num_id = str(next_num_id)
            next_num_id += 1
            num_map[old_num_id] = new_num_id

            cloned_num = deepcopy(num_el)
            cloned_num.set(f"{w_ns}numId", new_num_id)
            abs_val_el = cloned_num.find(f"{w_ns}abstractNumId")
            if abs_val_el is not None:
                old_ref = abs_val_el.get(f"{w_ns}val")
                if old_ref in abstract_map:
                    abs_val_el.set(f"{w_ns}val", abstract_map[old_ref])
            dst_num_el.append(cloned_num)

        # Обновляем ссылки numId в элементах тела поддокумента
        for num_id_el in num_id_els:
            old_val = num_id_el.get(f"{w_ns}val")
            if old_val in num_map:
                num_id_el.set(f"{w_ns}val", num_map[old_val])

    def _copy_element_relationships_native(
        self,
        src_doc: DocxDocument,
        element: etree._Element,
    ) -> None:
        """
        Переносит все связанные ресурсы (изображения, гиперссылки, диаграммы, OLE-объекты) для элемента.

        :param src_doc: Исходный поддокумент.
        :param element: XML-элемент тела поддокумента.
        """

        rel_attrs = (
            f"{{{REL_NS}}}embed",
            f"{{{REL_NS}}}id",
            f"{{{REL_NS}}}link",
            f"{{{REL_NS}}}dm",
            f"{{{REL_NS}}}lo",
            f"{{{REL_NS}}}qs",
            f"{{{REL_NS}}}cs",
        )

        rid_map: dict[str, str] = {}

        for node in element.iter():
            for attr in rel_attrs:
                old_rid = node.get(attr)
                if not old_rid or old_rid not in src_doc.part.rels:
                    continue

                if old_rid not in rid_map:
                    rel = src_doc.part.rels[old_rid]
                    if rel.is_external:
                        new_rid = self.doc.part.relate_to(
                            rel.target_ref,
                            rel.reltype,
                            is_external=True,
                        )
                    else:
                        target_part = rel.target_part
                        if rel.reltype == RT.IMAGE:
                            new_rid = self._relate_image_part(target_part)
                        else:
                            new_rid = self.doc.part.relate_to(target_part, rel.reltype)
                    rid_map[old_rid] = new_rid

                node.set(attr, rid_map[old_rid])

    def _relate_image_part(self, src_image_part: Any) -> str:
        """
        Привязывает часть изображения из поддокумента к основному документу без конфликта имён.

        :param src_image_part: Часть изображения из поддокумента.

        :return: Новый идентификатор связи rId в основном документе.
        """

        dst_package = self.doc.part.package
        if dst_package is not None and hasattr(dst_package, "parts"):
            existing_names = {str(p.partname) for p in dst_package.parts}
            partname = str(src_image_part.partname)
            if partname in existing_names:
                ext = partname.rsplit(".", 1)[-1] if "." in partname else "png"
                idx = 1
                while f"/word/media/subdoc_image{idx}.{ext}" in existing_names:
                    idx += 1
                src_image_part._partname = PackURI(f"/word/media/subdoc_image{idx}.{ext}")

        return self.doc.part.relate_to(src_image_part, RT.IMAGE)

    @staticmethod
    def _remove_header_footer_refs_native(element: etree._Element) -> None:
        """
        Удаляет ссылки на верхние и нижние колонтитулы из внутриабзацных sectPr поддокумента.

        :param element: XML-элемент поддокумента.
        """

        w_ns = f"{{{WORD_NS}}}"
        for sect_pr in element.iter(f"{w_ns}sectPr"):
            for tag in ("headerReference", "footerReference"):
                for ref in list(sect_pr.findall(f"{w_ns}{tag}")):
                    sect_pr.remove(ref)

    def _renumber_ids_native(self, src_doc: DocxDocument) -> None:
        """
        Перенумеровывает закладки и идентификаторы графических объектов в поддокументе.

        :param src_doc: Исходный поддокумент.
        """

        w_ns = f"{{{WORD_NS}}}"
        existing_bm_ids = [
            int(el.get(f"{w_ns}id", "0"))
            for el in self.doc.element.body.iter(f"{w_ns}bookmarkStart")
            if el.get(f"{w_ns}id", "").isdigit()
        ]
        next_bm_id = (max(existing_bm_ids) + 1) if existing_bm_ids else 100

        bm_map: dict[str, str] = {}
        for bm_start in src_doc.element.body.iter(f"{w_ns}bookmarkStart"):
            old_id = bm_start.get(f"{w_ns}id")
            if old_id is not None:
                new_id = str(next_bm_id)
                next_bm_id += 1
                bm_map[old_id] = new_id
                bm_start.set(f"{w_ns}id", new_id)

        for bm_end in src_doc.element.body.iter(f"{w_ns}bookmarkEnd"):
            old_id = bm_end.get(f"{w_ns}id")
            if old_id in bm_map:
                bm_end.set(f"{w_ns}id", bm_map[old_id])

        # Перенумеровываем wp:docPr и pic:cNvPr
        next_obj_id = 50000
        for doc_pr in src_doc.element.body.iter(f"{{{WP_NS}}}docPr"):
            doc_pr.set("id", str(next_obj_id))
            next_obj_id += 1
        for cnv_pr in src_doc.element.body.iter(f"{{{PIC_NS}}}cNvPr"):
            cnv_pr.set("id", str(next_obj_id))
            next_obj_id += 1

    def add_diagrams(self, doc: DocxDocument, element: etree._Element) -> None:
        """
        Переносит связи диаграмм SmartArt из поддокумента в основной документ.

        :param doc: Исходный поддокумент.
        :param element: XML-элемент, содержащий диаграмму.
        """

        if HAS_DOCXCOMPOSE:
            dgm_rels = _docxcompose_xpath(element, ".//dgm:relIds[@r:dm]")
            r_ns = _DOCXCOMPOSE_NS["r"]
        else:
            dgm_rels = element.xpath(
                ".//dgm:relIds[@r:dm]",
                namespaces={"dgm": DGM_NS, "r": REL_NS},
            )
            r_ns = REL_NS

        for dgm_rel in dgm_rels:
            for item, rt_type in (
                ("dm", RT.DIAGRAM_DATA),
                ("lo", RT.DIAGRAM_LAYOUT),
                ("qs", RT.DIAGRAM_QUICK_STYLE),
                ("cs", RT.DIAGRAM_COLORS),
            ):
                dm_rid = dgm_rel.get(f"{{{r_ns}}}{item}")
                if dm_rid and dm_rid in doc.part.rels:
                    dm_part = doc.part.rels[dm_rid].target_part
                    new_rid = self.doc.part.relate_to(dm_part, rt_type)
                    dgm_rel.set(f"{{{r_ns}}}{item}", new_rid)


class Subdoc:
    """ Класс поддокумента для динамического формирования или вставки внешнего файла .docx в основной шаблон. """

    def __init__(
        self,
        tpl: DocxTemplate,
        docpath: str | PathLike[str] | IO[bytes] | DocxDocument | None = None,
        remove_trailing_paragraph: bool = False,
    ) -> None:
        """
        Инициализация поддокумента.

        :param tpl: Экземпляр родительского шаблона DocxTemplate.
        :param docpath: Путь к существующему файлу .docx, файлоподобный объект, объект Document или None.
        :param remove_trailing_paragraph: Удалять ли пустой завершающий абзац при вставке.
        """

        self.tpl = tpl
        self.docx = tpl.get_docx()
        self.remove_trailing_paragraph = remove_trailing_paragraph

        # Поддержка передачи готового объекта docx.document.Document в памяти (Issue #603)
        if isinstance(docpath, DocxDocument) or (
            docpath is not None and hasattr(docpath, "element") and hasattr(docpath, "part")
        ):
            self.subdocx: DocxDocument = docpath  # type: ignore[assignment]
            has_external_doc = True
        else:
            self.subdocx = Document(docpath)
            has_external_doc = docpath is not None

        if has_external_doc:
            # Переносим пространства имён XML из поддокумента в основной документ (Issue #621: Pandoc DOCX)
            self._merge_namespaces()
            compose = SubdocComposer(self.docx)
            compose.attach_parts(self.subdocx)
        else:
            self.subdocx._part = self.docx._part

    def _merge_namespaces(self) -> None:
        """
        Регистрирует пространства имён XML из поддокумента в реестре python-docx и корневом элементе шаблона.
        """

        src_nsmap = getattr(self.subdocx.element, "nsmap", None) or {}
        dst_nsmap = getattr(self.docx._element, "nsmap", None) or {}

        for prefix, uri in src_nsmap.items():
            if not prefix or not uri:
                continue
            if prefix not in docx.oxml.ns.nsmap:
                docx.oxml.ns.nsmap[prefix] = uri
                docx.oxml.ns.pfxmap[uri] = prefix
                try:
                    etree.register_namespace(prefix, uri)
                except ValueError:
                    pass
            if prefix not in dst_nsmap:
                try:
                    attr_qname = f"{{{uri}}}__ns_init__"
                    self.docx._element.set(attr_qname, "1")
                    del self.docx._element.attrib[attr_qname]
                except Exception:
                    pass

    def __getattr__(self, name: str) -> Any:
        """
        Делегирует обращение к методам и атрибутам внутреннему объекту python-docx Document.

        :param name: Имя запрашиваемого атрибута.

        :return: Значение атрибута из self.subdocx.
        """

        return getattr(self.subdocx, name)

    def _get_xml(self) -> str:
        """
        Формирует XML-строку содержимого тела поддокумента для подстановки через Jinja2.

        :return: Строка XML без обрамляющего тега <w:body>.
        """

        body = self.subdocx.element.body
        if body.sectPr is not None:
            body.remove(body.sectPr)

        # Опциональное удаление пустого завершающего абзаца (Issue #629)
        if self.remove_trailing_paragraph:
            w_p = f"{{{WORD_NS}}}p"
            w_t = f"{{{WORD_NS}}}t"
            w_drawing = f"{{{WORD_NS}}}drawing"
            paragraphs = body.findall(w_p)
            if paragraphs:
                last_p = paragraphs[-1]
                has_text = any((t.text or "") for t in last_p.iter(w_t))
                has_drawing = next(last_p.iter(w_drawing), None) is not None
                if not has_text and not has_drawing:
                    body.remove(last_p)

        xml = re.sub(
            r"</?w:body[^>]*>",
            "",
            etree.tostring(body, encoding="unicode", pretty_print=False),
        )
        return xml

    def __bool__(self) -> bool:
        """
        Проверяет, содержит ли поддокумент какие-либо элементы.

        :return: True, если поддокумент не пуст, иначе False.
        """

        return len(self.subdocx.element.body) > 0

    def __unicode__(self) -> str:
        """
        Возвращает XML-представление поддокумента (для обратной совместимости).

        :return: XML-строка поддокумента.
        """

        return self._get_xml()

    def __str__(self) -> str:
        """
        Возвращает строковое XML-представление поддокумента.

        :return: XML-строка поддокумента.
        """

        return self._get_xml()

    def __html__(self) -> str:
        """
        Возвращает XML-представление поддокумента как безопасную разметку для Jinja2 autoescape.

        :return: XML-строка поддокумента.
        """

        return self._get_xml()

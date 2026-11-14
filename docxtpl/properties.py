# -*- coding: utf-8 -*-
"""
Модуль для работы со свойствами документа Word (Core, Extended/App и Custom свойства).
"""
from __future__ import annotations

import datetime as dt
import re
from collections.abc import Iterator, MutableMapping
from typing import TYPE_CHECKING, Any

from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import Part, XmlPart
from lxml import etree

if TYPE_CHECKING:
    from docx.document import Document as DocxDocument
    from jinja2 import Environment


# Пространства имён OpenXML для пользовательских и расширенных свойств документа
CUSTOM_PROPS_NS = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
VT_NS = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
APP_PROPS_NS = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# Стандартный FMTID для пользовательских свойств MS Office
FMTID_CUSTOM = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"

NSMAP_CUSTOM = {
    None: CUSTOM_PROPS_NS,
    "vt": VT_NS,
}

# Список всех строковых стандартных свойств (CoreProperties)
CORE_STRING_PROPERTIES: tuple[str, ...] = (
    "author",
    "category",
    "comments",
    "content_status",
    "identifier",
    "keywords",
    "language",
    "last_modified_by",
    "subject",
    "title",
    "version",
)

# Соответствие имён свойств CoreProperties именам XML-атрибутов в CT_CoreProperties
CORE_PROP_TO_ELEMENT: dict[str, str] = {
    "author": "creator",
    "category": "category",
    "comments": "description",
    "content_status": "contentStatus",
    "identifier": "identifier",
    "keywords": "keywords",
    "language": "language",
    "last_modified_by": "lastModifiedBy",
    "subject": "subject",
    "title": "title",
    "version": "version",
}

# Строковые свойства приложения (docProps/app.xml)
APP_STRING_PROPERTIES: tuple[str, ...] = (
    "Application",
    "AppVersion",
    "Company",
    "HyperlinkBase",
    "Manager",
    "Template",
)


class CustomProperties(MutableMapping[str, Any]):
    """ Менеджер пользовательских свойств документа Word (docProps/custom.xml). """

    def __init__(self, doc: DocxDocument) -> None:
        """
        Инициализация менеджера пользовательских свойств.

        :param doc: Экземпляр документа python-docx.
        """

        self._doc = doc
        self._part: Part | None = None
        self._element: etree._Element | None = None
        self._load_part()

    def _load_part(self) -> None:
        """
        Загружает существующую часть docProps/custom.xml из пакета документа, если она есть.
        """

        package = self._doc.part.package
        if package is None:
            return

        # Ищем связь на custom-properties в корневых связях пакета
        for rel in package.rels.values():
            if rel.reltype == RT.CUSTOM_PROPERTIES and not rel.is_external:
                self._part = rel.target_part
                break

        # Если связь не найдена, проверяем наличие части по URI
        if self._part is None:
            for part in package.parts:
                if str(part.partname) == "/docProps/custom.xml" or (
                    part.content_type == CT.OFC_CUSTOM_PROPERTIES
                ):
                    self._part = part
                    break

        if self._part is not None:
            if isinstance(self._part, XmlPart) and hasattr(self._part, "_element") and self._part._element is not None:
                # Перепарсим в стандартный lxml элемент для единообразной работы с пространствами имён
                blob = etree.tostring(self._part._element)
                self._element = etree.fromstring(blob)
            elif self._part.blob:
                self._element = etree.fromstring(self._part.blob)

    def _ensure_part(self) -> etree._Element:
        """
        Гарантирует существование XML-дерева и части docProps/custom.xml в пакете.

        :return: Корневой XML-элемент <Properties>.
        """

        if self._element is None:
            self._element = etree.Element(f"{{{CUSTOM_PROPS_NS}}}Properties", nsmap=NSMAP_CUSTOM)

        if self._part is None:
            package = self._doc.part.package
            blob = self._serialize_xml()
            self._part = Part(
                PackURI("/docProps/custom.xml"),
                CT.OFC_CUSTOM_PROPERTIES,
                blob,
                package,
            )
            if package is not None:
                package.relate_to(self._part, RT.CUSTOM_PROPERTIES)

        return self._element

    def _serialize_xml(self) -> bytes:
        """
        Сериализует текущее XML-дерево пользовательских свойств в байты UTF-8.

        :return: Байтовое представление XML-документа.
        """

        if self._element is None:
            return b""
        return etree.tostring(
            self._element,
            encoding="UTF-8",
            xml_declaration=True,
            standalone="yes",
        )

    def sync(self) -> None:
        """
        Синхронизирует изменения XML-дерева обратно в часть пакета docx.
        """

        if self._element is None or self._part is None:
            return

        # Перенумеровываем pid (в OpenXML pid обязан начинаться с 2 и идти по порядку)
        for idx, prop_el in enumerate(
            self._element.findall(f"{{{CUSTOM_PROPS_NS}}}property"), start=2
        ):
            prop_el.set("pid", str(idx))

        blob = self._serialize_xml()
        self._part._blob = blob
        if isinstance(self._part, XmlPart) and hasattr(self._part, "_element"):
            try:
                from docx.opc.oxml import parse_xml

                self._part._element = parse_xml(blob)
            except Exception:
                pass

    def _find_property_element(self, name: str) -> etree._Element | None:
        """
        Находит XML-элемент <property> по имени свойства.

        :param name: Имя пользовательского свойства.

        :return: Найденный XML-элемент или None.
        """

        if self._element is None:
            return None

        for prop_el in self._element.findall(f"{{{CUSTOM_PROPS_NS}}}property"):
            if prop_el.get("name") == name:
                return prop_el
        return None

    @staticmethod
    def _decode_value(prop_el: etree._Element) -> Any:
        """
        Декодирует значение свойства из дочернего элемента пространства имён vt.

        :param prop_el: XML-элемент <property>.

        :return: Преобразованное значение Python (str, int, float, bool, datetime).
        """

        if len(prop_el) == 0:
            return ""

        val_el = prop_el[0]
        local_name = etree.QName(val_el).localname
        raw_text = val_el.text or ""

        if local_name in ("lpwstr", "lpstr", "bstr"):
            return raw_text
        if local_name == "bool":
            return raw_text.strip().lower() in ("true", "1")
        if local_name in ("i1", "i2", "i4", "i8", "int", "ui1", "ui2", "ui4", "ui8", "uint"):
            try:
                return int(raw_text.strip())
            except ValueError:
                return 0
        if local_name in ("r4", "r8", "decimal"):
            try:
                return float(raw_text.strip())
            except ValueError:
                return 0.0
        if local_name in ("filetime", "date"):
            text = raw_text.strip()
            try:
                if text.endswith("Z"):
                    text = text[:-1] + "+00:00"
                return dt.datetime.fromisoformat(text)
            except ValueError:
                return raw_text

        return raw_text

    @staticmethod
    def _encode_value(prop_el: etree._Element, value: Any) -> None:
        """
        Записывает значение Python в XML-элемент <property> с соответствующим типом vt.

        :param prop_el: XML-элемент <property>.
        :param value: Значение для записи.
        """

        # Очищаем старое содержимое элемента свойства
        for child in list(prop_el):
            prop_el.remove(child)

        # Важно: bool проверяется раньше int, так как в Python bool является подклассом int
        if isinstance(value, bool):
            val_el = etree.SubElement(prop_el, f"{{{VT_NS}}}bool")
            val_el.text = "true" if value else "false"
        elif isinstance(value, int):
            tag = "i8" if (value < -2147483648 or value > 2147483647) else "i4"
            val_el = etree.SubElement(prop_el, f"{{{VT_NS}}}{tag}")
            val_el.text = str(value)
        elif isinstance(value, float):
            val_el = etree.SubElement(prop_el, f"{{{VT_NS}}}r8")
            val_el.text = str(value)
        elif isinstance(value, dt.datetime):
            val_el = etree.SubElement(prop_el, f"{{{VT_NS}}}filetime")
            if value.tzinfo is not None:
                utc_val = value.astimezone(dt.timezone.utc)
            else:
                utc_val = value
            val_el.text = utc_val.strftime("%Y-%m-%dT%H:%M:%SZ")
        elif isinstance(value, dt.date):
            val_el = etree.SubElement(prop_el, f"{{{VT_NS}}}filetime")
            val_el.text = f"{value.strftime('%Y-%m-%d')}T00:00:00Z"
        else:
            val_el = etree.SubElement(prop_el, f"{{{VT_NS}}}lpwstr")
            val_el.text = "" if value is None else str(value)

    def __getitem__(self, name: str) -> Any:
        """
        Возвращает значение пользовательского свойства по его имени.

        :param name: Имя свойства.

        :return: Значение свойства.
        """

        prop_el = self._find_property_element(name)
        if prop_el is None:
            raise KeyError(f"Пользовательское свойство '{name}' не найдено.")
        return self._decode_value(prop_el)

    def __setitem__(self, name: str, value: Any) -> None:
        """
        Устанавливает или обновляет значение пользовательского свойства.

        :param name: Имя свойства.
        :param value: Новое значение свойства.
        """

        root = self._ensure_part()
        prop_el = self._find_property_element(name)
        if prop_el is None:
            existing_count = len(root.findall(f"{{{CUSTOM_PROPS_NS}}}property"))
            prop_el = etree.SubElement(
                root,
                f"{{{CUSTOM_PROPS_NS}}}property",
                fmtid=FMTID_CUSTOM,
                pid=str(existing_count + 2),
                name=str(name),
            )
        self._encode_value(prop_el, value)
        self.sync()

    def __delitem__(self, name: str) -> None:
        """
        Удаляет пользовательское свойство по имени.

        :param name: Имя свойства для удаления.
        """

        prop_el = self._find_property_element(name)
        if prop_el is None or self._element is None:
            raise KeyError(f"Пользовательское свойство '{name}' не найдено.")
        self._element.remove(prop_el)
        self.sync()

    def __iter__(self) -> Iterator[str]:
        """
        Возвращает итератор по именам всех пользовательских свойств.

        :return: Итератор строковых имён свойств.
        """

        if self._element is None:
            return iter(())
        names = [
            prop_el.get("name", "")
            for prop_el in self._element.findall(f"{{{CUSTOM_PROPS_NS}}}property")
            if prop_el.get("name") is not None
        ]
        return iter(names)

    def __len__(self) -> int:
        """
        Возвращает количество пользовательских свойств.

        :return: Число свойств.
        """

        if self._element is None:
            return 0
        return len(self._element.findall(f"{{{CUSTOM_PROPS_NS}}}property"))

    def __repr__(self) -> str:
        """
        Возвращает строковое представление всех пользовательских свойств.

        :return: Строка вида CustomProperties({...}).
        """

        return f"CustomProperties({self.to_dict()!r})"

    def set(self, name: str, value: Any) -> None:
        """
        Устанавливает значение пользовательского свойства.

        :param name: Имя свойства.
        :param value: Значение свойства.
        """

        self[name] = value

    def delete(self, name: str) -> bool:
        """
        Удаляет пользовательское свойство, если оно существует.

        :param name: Имя свойства.

        :return: True, если свойство было удалено, иначе False.
        """

        if name in self:
            del self[name]
            return True
        return False

    def to_dict(self) -> dict[str, Any]:
        """
        Преобразует все пользовательские свойства в обычный словарь Python.

        :return: Словарь {имя_свойства: значение}.
        """

        return {k: self[k] for k in self}

    def render_properties(self, context: dict[str, Any], jinja_env: Environment) -> None:
        """
        Рендерит строковые пользовательские свойства как шаблоны Jinja2.

        :param context: Словарь контекста для рендеринга.
        :param jinja_env: Окружение Jinja2.
        """

        if self._element is None:
            return

        changed = False
        for prop_el in self._element.findall(f"{{{CUSTOM_PROPS_NS}}}property"):
            if len(prop_el) == 0:
                continue
            val_el = prop_el[0]
            local_name = etree.QName(val_el).localname
            if local_name in ("lpwstr", "lpstr", "bstr") and val_el.text:
                initial = val_el.text
                if "{{" in initial or "{%" in initial or "{#" in initial:
                    template = jinja_env.from_string(initial)
                    rendered = template.render(context)
                    if rendered != initial:
                        val_el.text = rendered
                        changed = True

        if changed:
            self.sync()

    @staticmethod
    def _extract_docproperty_name(instr: str) -> str | None:
        """
        Извлекает имя свойства из инструкции поля Word DOCPROPERTY.

        :param instr: Строка инструкции поля (например, ' DOCPROPERTY "MyProp" \\* MERGEFORMAT ').

        :return: Имя свойства или None, если это не поле DOCPROPERTY.
        """

        match = re.search(
            r'^\s*DOCPROPERTY\s+(?:"([^"]+)"|([^\s\\"]+))',
            instr,
            flags=re.IGNORECASE,
        )
        if not match:
            return None
        return match.group(1) or match.group(2)

    def update_fields_in_element(
        self,
        root_element: etree._Element,
        extra_properties: dict[str, Any] | None = None,
    ) -> None:
        """
        Обновляет отображаемые значения полей DOCPROPERTY внутри переданного XML-дерева.

        :param root_element: Корневой XML-элемент (например, body или колонтитул).
        :param extra_properties: Дополнительные свойства (например, Core/App свойства) для подстановки.
        """

        props = self.to_dict()
        if extra_properties:
            merged = {**extra_properties, **props}
        else:
            merged = props

        if not merged:
            return

        # Собираем карту с учётом регистронезависимого поиска (в Word имена DOCPROPERTY регистронезависимы)
        lower_map: dict[str, str] = {}
        for k, v in merged.items():
            if v is None:
                continue
            if isinstance(v, bool):
                str_val = "True" if v else "False"
            elif isinstance(v, (dt.datetime, dt.date)):
                str_val = v.strftime("%Y-%m-%d")
            else:
                str_val = str(v)
            lower_map[k.lower()] = str_val

        w_ns = f"{{{WORD_NS}}}"

        # 1. Простые поля <w:fldSimple w:instr="DOCPROPERTY ...">
        for fld_simple in root_element.iter(f"{w_ns}fldSimple"):
            instr = fld_simple.get(f"{w_ns}instr") or fld_simple.get("instr") or ""
            prop_name = self._extract_docproperty_name(instr)
            if prop_name and prop_name.lower() in lower_map:
                new_val = lower_map[prop_name.lower()]
                t_elements = list(fld_simple.iter(f"{w_ns}t"))
                if t_elements:
                    t_elements[0].text = new_val
                    for extra_t in t_elements[1:]:
                        extra_t.text = ""

        # 2. Сложные поля (<w:fldChar w:fldCharType="begin"/> ... <w:instrText> ... <w:fldChar w:fldCharType="separate"/> ... <w:fldChar w:fldCharType="end"/>)
        for p_el in root_element.iter(f"{w_ns}p"):
            in_field = False
            in_separate = False
            instr_parts: list[str] = []
            value_t_elements: list[etree._Element] = []

            for r_el in p_el.findall(f"{w_ns}r"):
                fld_char = r_el.find(f"{w_ns}fldChar")
                if fld_char is not None:
                    fld_type = fld_char.get(f"{w_ns}fldCharType") or fld_char.get("fldCharType")
                    if fld_type == "begin":
                        in_field = True
                        in_separate = False
                        instr_parts = []
                        value_t_elements = []
                    elif fld_type == "separate" and in_field:
                        in_separate = True
                    elif fld_type == "end" and in_field:
                        prop_name = self._extract_docproperty_name("".join(instr_parts))
                        if prop_name and prop_name.lower() in lower_map and value_t_elements:
                            value_t_elements[0].text = lower_map[prop_name.lower()]
                            for extra_t in value_t_elements[1:]:
                                extra_t.text = ""
                        in_field = False
                        in_separate = False
                        instr_parts = []
                        value_t_elements = []
                    continue

                if in_field and not in_separate:
                    for instr_el in r_el.findall(f"{w_ns}instrText"):
                        if instr_el.text:
                            instr_parts.append(instr_el.text)
                elif in_field and in_separate:
                    for t_el in r_el.findall(f"{w_ns}t"):
                        value_t_elements.append(t_el)

    def dissolve_fields(self, name: str | None = None) -> None:
        """
        Заменяет поля DOCPROPERTY в документе их текущими текстовыми значениями (удаляя саму конструкцию поля).

        :param name: Имя конкретного свойства (если None — заменяются все поля пользовательских свойств).
        """

        w_ns = f"{{{WORD_NS}}}"
        target_name = name.lower() if name is not None else None
        body = self._doc.element.body

        # Обрабатываем простые поля <w:fldSimple>
        for fld_simple in list(body.iter(f"{w_ns}fldSimple")):
            instr = fld_simple.get(f"{w_ns}instr") or fld_simple.get("instr") or ""
            prop_name = self._extract_docproperty_name(instr)
            if prop_name is None:
                continue
            if target_name is not None and prop_name.lower() != target_name:
                continue

            parent = fld_simple.getparent()
            if parent is None:
                continue
            idx = parent.index(fld_simple)
            for child in list(fld_simple):
                parent.insert(idx, child)
                idx += 1
            parent.remove(fld_simple)


def get_core_properties_dict(doc: DocxDocument) -> dict[str, Any]:
    """
    Возвращает словарь со всеми основными свойствами документа (docProps/core.xml).

    :param doc: Экземпляр документа python-docx.

    :return: Словарь основных свойств документа.
    """

    cp = doc.core_properties
    result: dict[str, Any] = {}
    for prop in CORE_STRING_PROPERTIES:
        result[prop] = getattr(cp, prop, "") or ""

    for dt_prop in ("created", "modified", "last_printed"):
        result[dt_prop] = getattr(cp, dt_prop, None)

    result["revision"] = getattr(cp, "revision", 0)
    return result


def set_core_property_value(doc: DocxDocument, name: str, value: Any) -> None:
    """
    Устанавливает значение стандартного свойства документа (включая строки длиннее 255 символов).

    :param doc: Экземпляр документа python-docx.
    :param name: Имя свойства CoreProperties (например, 'title', 'author', 'comments').
    :param value: Значение свойства.
    """

    cp = doc.core_properties
    if name in CORE_PROP_TO_ELEMENT:
        str_val = "" if value is None else str(value)
        if len(str_val) <= 255:
            setattr(cp, name, str_val)
        else:
            # Обходим ограничение python-docx в 255 символов, записывая напрямую в XML-элемент
            el_name = CORE_PROP_TO_ELEMENT[name]
            el = cp._element._get_or_add(el_name)
            el.text = str_val
    elif name in ("created", "modified", "last_printed"):
        if isinstance(value, dt.datetime):
            setattr(cp, name, value)
    elif name == "revision":
        try:
            cp.revision = max(1, int(value))
        except (ValueError, TypeError):
            pass
    else:
        raise ValueError(f"Неизвестное стандартное свойство документа: '{name}'.")


def get_app_properties_dict(doc: DocxDocument) -> dict[str, str]:
    """
    Возвращает словарь строковых расширенных свойств приложения из docProps/app.xml.

    :param doc: Экземпляр документа python-docx.

    :return: Словарь расширенных свойств (Company, Manager, Application и др.).
    """

    part = _get_app_properties_part(doc)
    if part is None or not part.blob:
        return {}

    root = etree.fromstring(part.blob)
    result: dict[str, str] = {}
    for prop in APP_STRING_PROPERTIES:
        el = root.find(f"{{{APP_PROPS_NS}}}{prop}")
        if el is not None and el.text is not None:
            result[prop] = el.text
    return result


def set_app_properties(doc: DocxDocument, properties: dict[str, Any]) -> None:
    """
    Устанавливает расширенные свойства приложения в docProps/app.xml (например, Company, Manager).

    :param doc: Экземпляр документа python-docx.
    :param properties: Словарь устанавливаемых свойств.
    """

    part = _get_app_properties_part(doc)
    if part is None or not part.blob:
        return

    root = etree.fromstring(part.blob)
    alias_map = {p.lower(): p for p in APP_STRING_PROPERTIES}
    alias_map["app_version"] = "AppVersion"
    alias_map["hyperlink_base"] = "HyperlinkBase"

    changed = False
    for key, val in properties.items():
        xml_name = alias_map.get(key.lower(), key)
        el = root.find(f"{{{APP_PROPS_NS}}}{xml_name}")
        if el is None:
            el = etree.SubElement(root, f"{{{APP_PROPS_NS}}}{xml_name}")
        el.text = "" if val is None else str(val)
        changed = True

    if changed:
        part._blob = etree.tostring(
            root, encoding="UTF-8", xml_declaration=True, standalone="yes"
        )


def render_app_properties(
    doc: DocxDocument,
    context: dict[str, Any],
    jinja_env: Environment,
) -> None:
    """
    Рендерит шаблоны Jinja2 внутри строковых свойств docProps/app.xml.

    :param doc: Экземпляр документа python-docx.
    :param context: Контекст переменных Jinja2.
    :param jinja_env: Окружение Jinja2.
    """

    part = _get_app_properties_part(doc)
    if part is None or not part.blob:
        return

    root = etree.fromstring(part.blob)
    changed = False
    for prop in APP_STRING_PROPERTIES:
        el = root.find(f"{{{APP_PROPS_NS}}}{prop}")
        if el is not None and el.text:
            initial = el.text
            if "{{" in initial or "{%" in initial or "{#" in initial:
                template = jinja_env.from_string(initial)
                rendered = template.render(context)
                if rendered != initial:
                    el.text = rendered
                    changed = True

    if changed:
        part._blob = etree.tostring(
            root, encoding="UTF-8", xml_declaration=True, standalone="yes"
        )


def _get_app_properties_part(doc: DocxDocument) -> Part | None:
    """
    Находит часть docProps/app.xml в пакете документа.

    :param doc: Экземпляр документа python-docx.

    :return: Часть пакета или None.
    """

    package = doc.part.package
    if package is None:
        return None

    for rel in package.rels.values():
        if rel.reltype == RT.EXTENDED_PROPERTIES and not rel.is_external:
            return rel.target_part

    for part in package.parts:
        if str(part.partname) == "/docProps/app.xml":
            return part

    return None

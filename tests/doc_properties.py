# -*- coding: utf-8 -*-
"""
Тестирование работы со стандартными (Core), расширенными (App) и пользовательскими (Custom) свойствами документа.
"""
from __future__ import annotations

import datetime as dt

from docxtpl import DocxTemplate

doctemplate = r"templates/doc_properties_tpl.docx"

tpl = DocxTemplate(doctemplate)

# Устанавливаем пользовательские свойства документа разных типов (включая шаблонные выражения Jinja2)
tpl.set_custom_property("НомерДоговора", "ДГ-{{ contract_num }}")
tpl.set_custom_property("СуммаДоговора", 150000.75)
tpl.set_custom_property("КоличествоЭтапов", 4)
tpl.set_custom_property("Подписан", True)
tpl.set_custom_property("ДатаДоговора", dt.date(2026, 11, 15))

# Устанавливаем расширенные свойства приложения (docProps/app.xml)
tpl.set_app_properties(Company="ООО «{{ company }}»", Manager="Иван Петров")

context = {
    "test": "HelloWorld",
    "contract_num": "2026/42",
    "company": "ТехноСофт",
}

tpl.render(context)
tpl.save("output/doc_properties.docx")

# Проверяем, что свойства корректно сохранились и отрендерились в выходном файле
verified = DocxTemplate("output/doc_properties.docx")
assert verified.get_core_properties()["title"] == "HelloWorld"
assert verified.custom_properties["НомерДоговора"] == "ДГ-2026/42"
assert verified.custom_properties["СуммаДоговора"] == 150000.75
assert verified.custom_properties["КоличествоЭтапов"] == 4
assert verified.custom_properties["Подписан"] is True
assert verified.get_app_properties()["Company"] == "ООО «ТехноСофт»"
assert verified.get_app_properties()["Manager"] == "Иван Петров"

# -*- coding: utf-8 -*-
"""
Комплексные тесты для новых возможностей и исправленных ошибок:
- Замена переменных в математических формулах Word OMML (<m:oMath>, <m:r><m:t>, <m:d>, <m:sSub>).
- Поддержка операторов % и символов } в управляющих тегах {%p %}, {%tr %}, {%tc %}, {%r %} (Issue #660).
- Очистка узких неразрывных пробелов \\u202f и \\u00a0 внутри тегов Jinja2 (Issue #628).
- Проверка пустых объектов RichText и RichTextParagraph в условиях {% if %} (Issue #601).
- Создание Subdoc из объекта python-docx Document в памяти (Issue #603) и удаление пустого абзаца (Issue #629).
- Безопасный режим песочницы Jinja2 sandbox=True (Issue #644).
- Асинхронный рендеринг render_async и save_async (Issue #611).
- Работа с пользовательскими свойствами Word и полями DOCPROPERTY.
"""
from __future__ import annotations

import io

from docx import Document
from docx.oxml import parse_xml
from jinja2.exceptions import SecurityError

from docxtpl import DocxTemplate, Listing, RichText, RichTextParagraph


def _build_formula_template_bytes() -> io.BytesIO:
    """
    Создаёт в памяти тестовый документ .docx со сложными математическими формулами OMML,
    разбитыми прогонами <m:r><m:t>, разделителями <m:d>, авто-индексами <m:sSub>,
    полями DOCPROPERTY и тегами {%p %} с оператором %.

    :return: Бинарный поток шаблона .docx.
    """

    doc = Document()

    # 1. Формула OMML, где '{{ mass_1 }}', '{{ acc }}' и '{{ r_dist }}' разбиты по разным <m:r><m:t>,
    # обернуты в <m:d> (скобки формул Word) и превращены в <m:sSub> из-за знака '_'
    p_math = doc.add_paragraph()
    omath_xml = """
    <m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
                 xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <m:oMath>
        <m:r><m:t>F=</m:t></m:r>
        <m:f>
          <m:num>
            <m:d>
              <m:dPr>
                <m:begChr m:val="{"/>
                <m:endChr m:val="}"/>
              </m:dPr>
              <m:e>
                <m:d>
                  <m:dPr>
                    <m:begChr m:val="{"/>
                    <m:endChr m:val="}"/>
                  </m:dPr>
                  <m:e>
                    <m:sSub>
                      <m:e><m:r><m:t>mass</m:t></m:r></m:e>
                      <m:sub><m:r><m:t>1</m:t></m:r></m:sub>
                    </m:sSub>
                  </m:e>
                </m:d>
              </m:e>
            </m:d>
            <m:r><m:t>⋅</m:t></m:r>
            <m:r><m:t>{</m:t></m:r>
            <m:r><m:t>{</m:t></m:r>
            <m:r><m:t>a</m:t></m:r>
            <m:r><m:t>c</m:t></m:r>
            <m:r><m:t>c</m:t></m:r>
            <m:r><m:t>}</m:t></m:r>
            <m:r><m:t>}</m:t></m:r>
          </m:num>
          <m:den>
            <m:sSub>
              <m:e><m:r><m:t>{{ r</m:t></m:r></m:e>
              <m:sub><m:r><m:t>dist }}</m:t></m:r></m:sub>
            </m:sSub>
          </m:den>
        </m:f>
      </m:oMath>
    </m:oMathPara>
    """
    p_math._p.append(parse_xml(omath_xml))

    # 2. Тег {%p %} с оператором остатка от деления % и словарём {} (Issue #660) + \u202f (Issue #628)
    doc.add_paragraph("{%p if (num % 2 == 0) and data.get('ok', {'k': 1}) \u202f%}")
    doc.add_paragraph("Чётное число: {{ num }}")
    doc.add_paragraph("{%p endif %}")

    # 3. Проверка пустого и непустого RichText в условии (Issue #601)
    doc.add_paragraph("{%p if empty_rt %}")
    doc.add_paragraph("Этот текст НЕ должен появиться")
    doc.add_paragraph("{%p endif %}")

    doc.add_paragraph("{%p if filled_rt %}")
    doc.add_paragraph("Форматированный текст: {{r filled_rt }}")
    doc.add_paragraph("{%p endif %}")

    # 4. Вставка поддокумента из памяти (Issue #603)
    doc.add_paragraph("{{p mem_subdoc }}")

    # 5. Поле DOCPROPERTY для пользовательского свойства
    p_field = doc.add_paragraph()
    fld_xml = (
        '<w:fldSimple xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'w:instr=\' DOCPROPERTY "Проект" \\* MERGEFORMAT \'>'
        "<w:r><w:t>СтароеЗначение</w:t></w:r>"
        "</w:fldSimple>"
    )
    p_field._p.append(parse_xml(fld_xml))

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


def test_all_features() -> None:
    """
    Запускает проверку формул, тегов с % и }, RichText bool, in-memory Subdoc, кастомных свойств и sandbox.
    """

    tpl_stream = _build_formula_template_bytes()
    tpl = DocxTemplate(tpl_stream)

    # Создаём документ в памяти и передаём его напрямую в new_subdoc (Issue #603)
    in_mem_doc = Document()
    in_mem_doc.add_paragraph("Абзац из поддокумента в памяти")
    in_mem_doc.add_paragraph("")
    sd = tpl.new_subdoc(in_mem_doc, remove_trailing_paragraph=True)

    tpl.custom_properties["Проект"] = "АИМ-{{ project_code }}"

    empty_rt = RichText()
    filled_rt = RichText("Важно", bold=True, color="ff0000")
    zero_rt = RichText(0)
    assert not bool(empty_rt)
    assert bool(filled_rt)
    assert bool(zero_rt)
    assert not bool(RichTextParagraph())
    assert bool(RichTextParagraph("Абзац"))
    assert not bool(Listing(""))
    assert bool(Listing("Текст"))

    context = {
        "mass_1": 10,
        "acc": 9.81,
        "r_dist": 2,
        "num": 4,
        "data": {"ok": True},
        "empty_rt": empty_rt,
        "filled_rt": filled_rt,
        "mem_subdoc": sd,
        "project_code": "2026",
    }

    # Проверяем валидацию шаблона и проверку полноты контекста до рендеринга
    assert tpl.is_template_valid()
    assert tpl.validate_template() == []

    incomplete_report = tpl.check_context({"mass_1": 10, "extra_unused": 123})
    assert not incomplete_report.is_valid
    assert not bool(incomplete_report)
    assert "acc" in incomplete_report.missing
    assert "project_code" in incomplete_report["missing"]
    assert "extra_unused" in incomplete_report.unused

    full_report = tpl.check_context(context)
    assert full_report.is_valid
    assert bool(full_report)
    assert len(full_report.missing) == 0
    assert len(full_report.unused) == 0

    # Проверяем включение автообновления полей при открытии в Word (<w:updateFields w:val="true"/>)
    assert not tpl.get_update_fields_on_open()
    tpl.update_fields_on_open(True)
    assert tpl.get_update_fields_on_open()

    # Проверяем асинхронный рендеринг и сохранение (Issue #611)
    out_stream = io.BytesIO()
    for coro in (tpl.render_async(context, update_fields=True), tpl.save_async(out_stream)):
        try:
            coro.send(None)
        except StopIteration:
            pass
    tpl.save("output/formulas_and_features.docx")

    # Проверяем, что настройка updateFields сохранилась в выходном файле
    reloaded_tpl = DocxTemplate("output/formulas_and_features.docx")
    assert reloaded_tpl.get_update_fields_on_open()

    # Проверяем содержимое сгенерированного документа
    out_xml = tpl.get_xml()
    assert "10" in out_xml
    assert "9.81" in out_xml
    assert ">2<" in out_xml
    assert "Чётное число: 4" in out_xml
    assert "Этот текст НЕ должен появиться" not in out_xml
    assert "Важно" in out_xml
    assert "Абзац из поддокумента в памяти" in out_xml
    assert "АИМ-2026" in out_xml
    assert tpl.custom_properties["Проект"] == "АИМ-2026"

    # Проверяем обнаружение синтаксических ошибок в нескольких частях шаблона через validate_template()
    broken_doc = Document()
    broken_doc.add_paragraph("Корректный абзац {{ ok_var }}")
    broken_doc.add_paragraph("Ошибочный абзац {% if unclosed %}")
    broken_buf = io.BytesIO()
    broken_doc.save(broken_buf)
    broken_buf.seek(0)

    broken_tpl = DocxTemplate(broken_buf)
    broken_tpl.set_custom_property("БитоеСвойство", "{{ bad_expr + }}")
    assert not broken_tpl.is_template_valid()
    val_errors = broken_tpl.validate_template()
    assert len(val_errors) == 2
    error_parts = {e.part_name for e in val_errors}
    assert "body" in error_parts
    assert "custom_properties.БитоеСвойство" in error_parts

    # Проверяем защиту SandboxedEnvironment от SSTI (Issue #644)
    evil_doc = Document()
    evil_doc.add_paragraph("{{ ''.__class__.__mro__[1] }}")
    evil_buf = io.BytesIO()
    evil_doc.save(evil_buf)
    evil_buf.seek(0)

    safe_tpl = DocxTemplate(evil_buf, sandbox=True)
    try:
        safe_tpl.render({})
        raise AssertionError("SandboxedEnvironment должен блокировать доступ к __class__!")
    except SecurityError:
        pass


if __name__ == "__main__":
    test_all_features()

# -*- coding: utf-8 -*-
"""
Интерфейс командной строки (CLI) для генерации файла .docx из шаблона и JSON-данных.
"""
from __future__ import annotations

import argparse
import json
import os
from typing import Any

from .template import DocxTemplate, TemplateError

TEMPLATE_ARG = "template_path"
JSON_ARG = "json_path"
OUTPUT_ARG = "output_filename"
OVERWRITE_ARG = "overwrite"
QUIET_ARG = "quiet"


def make_arg_parser() -> argparse.ArgumentParser:
    """
    Создаёт и настраивает парсер аргументов командной строки для запуска модуля docxtpl.

    :return: Экземпляр argparse.ArgumentParser.
    """

    parser = argparse.ArgumentParser(
        usage=f"python -m docxtpl [-h] [-o] [-q] {TEMPLATE_ARG} {JSON_ARG} {OUTPUT_ARG}",
        description="Создание документа .docx из существующего шаблона .docx и файла данных JSON.",
    )
    parser.add_argument(
        TEMPLATE_ARG,
        type=str,
        help="Путь к входному файлу шаблона .docx.",
    )
    parser.add_argument(
        JSON_ARG,
        type=str,
        help="Путь к файлу JSON с данными контекста.",
    )
    parser.add_argument(
        OUTPUT_ARG,
        type=str,
        help="Путь для сохранения сгенерированного файла .docx.",
    )
    parser.add_argument(
        "-" + OVERWRITE_ARG[0],
        "--" + OVERWRITE_ARG,
        action="store_true",
        help="Перезаписывать выходной файл без запроса подтверждения, если он уже существует.",
    )
    parser.add_argument(
        "-" + QUIET_ARG[0],
        "--" + QUIET_ARG,
        action="store_true",
        help="Не выводить информационные сообщения в консоль.",
    )
    return parser


def get_args(parser: argparse.ArgumentParser) -> dict[str, Any]:
    """
    Разбирает аргументы командной строки с помощью переданного парсера.

    :param parser: Экземпляр парсера аргументов.

    :return: Словарь разобранных аргументов.
    """

    try:
        parsed_args = vars(parser.parse_args())
        return parsed_args
    # Ошибки аргументов вызывают SystemExit с кодом 2, а флаг --help / -h — с кодом 0
    except SystemExit as e:
        if e.code == 0:
            raise SystemExit
        raise RuntimeError(f"Правильное использование:\n{parser.usage}")


def is_argument_valid(arg_name: str, arg_value: Any, overwrite: bool) -> bool:
    """
    Проверяет корректность значения отдельного аргумента командной строки.

    :param arg_name: Имя проверяемого аргумента.
    :param arg_value: Значение аргумента.
    :param overwrite: Флаг принудительной перезаписи выходного файла.

    :return: True, если аргумент корректен, иначе False.
    """

    # Базовые проверки существования файлов и расширений
    if arg_name == TEMPLATE_ARG:
        return isinstance(arg_value, str) and os.path.isfile(arg_value) and arg_value.endswith(".docx")
    if arg_name == JSON_ARG:
        return isinstance(arg_value, str) and os.path.isfile(arg_value) and arg_value.endswith(".json")
    if arg_name == OUTPUT_ARG:
        return isinstance(arg_value, str) and arg_value.endswith(".docx") and check_exists_ask_overwrite(
            arg_value, overwrite
        )
    if arg_name in [OVERWRITE_ARG, QUIET_ARG]:
        return arg_value in [True, False]
    return False


def check_exists_ask_overwrite(arg_value: str, overwrite: bool) -> bool:
    """
    Проверяет существование выходного файла и при необходимости запрашивает подтверждение перезаписи.

    :param arg_value: Путь к выходному файлу.
    :param overwrite: Разрешена ли перезапись без подтверждения.

    :return: True, если запись разрешена.
    """

    if os.path.exists(arg_value) and not overwrite:
        try:
            msg = (
                f"Файл {arg_value} уже существует, перезаписать существующий файл? (y/n): "
            )
            if input(msg).strip().lower() in ("y", "д", "yes", "да"):
                return True
            raise OSError
        except OSError:
            raise RuntimeError(
                f"Файл {arg_value} уже существует, пожалуйста, выберите другое имя."
            )
    return True


def validate_all_args(parsed_args: dict[str, Any]) -> None:
    """
    Проверяет валидность всех разобранных аргументов командной строки.

    :param parsed_args: Словарь аргументов командной строки.
    """

    overwrite = bool(parsed_args[OVERWRITE_ARG])
    arg_name = ""
    arg_value: Any = None
    try:
        for arg_name, arg_value in parsed_args.items():
            if not is_argument_valid(arg_name, arg_value, overwrite):
                raise AssertionError
    except AssertionError:
        raise RuntimeError(
            f'Указанный аргумент {arg_name} "{arg_value}" некорректен.'
        )


def get_json_data(json_path: str) -> dict[str, Any]:
    """
    Загружает данные контекста из JSON-файла в кодировке UTF-8 (Issue #598).

    :param json_path: Абсолютный или относительный путь к файлу JSON.

    :return: Словарь данных, загруженный из JSON.
    """

    with open(json_path, encoding="utf-8") as file:
        try:
            json_data: dict[str, Any] = json.load(file)
            return json_data
        except json.JSONDecodeError as e:
            print(
                f"Ошибка разбора JSON в строке {e.lineno}, столбце {e.colno} "
                f"при чтении файла {json_path}"
            )
            raise RuntimeError("Не удалось загрузить данные из JSON-файла.")


def make_docxtemplate(template_path: str) -> DocxTemplate:
    """
    Создаёт экземпляр шаблона DocxTemplate по указанному пути.

    :param template_path: Путь к файлу шаблона .docx.

    :return: Экземпляр DocxTemplate.
    """

    try:
        return DocxTemplate(template_path)
    except TemplateError:
        raise RuntimeError("Не удалось инициализировать шаблон docx.")


def render_docx(doc: DocxTemplate, json_data: dict[str, Any]) -> DocxTemplate:
    """
    Выполняет рендеринг документа DocxTemplate с переданными JSON-данными.

    :param doc: Экземпляр шаблона DocxTemplate.
    :param json_data: Словарь контекста шаблона.

    :return: Отрендеренный экземпляр DocxTemplate.
    """

    try:
        doc.render(json_data)
        return doc
    except TemplateError:
        raise RuntimeError("Произошла ошибка при рендеринге документа docx.")


def save_file(doc: DocxTemplate, parsed_args: dict[str, Any]) -> None:
    """
    Сохраняет отрендеренный документ по пути, указанному в аргументах командной строки.

    :param doc: Отрендеренный экземпляр DocxTemplate.
    :param parsed_args: Словарь аргументов командной строки.
    """

    try:
        output_path = parsed_args[OUTPUT_ARG]
        doc.save(output_path)
        if not parsed_args[QUIET_ARG]:
            print(f"Документ успешно сгенерирован и сохранён по пути: {output_path}")
    except OSError as e:
        print(f"{e.strerror}. Не удалось сохранить файл {e.filename}.")
        raise RuntimeError("Ошибка сохранения выходного файла.")


def main() -> None:
    """
    Точка входа для запуска генерации документа из командной строки (python -m docxtpl).
    """

    parser = make_arg_parser()
    parsed_args = get_args(parser)
    try:
        validate_all_args(parsed_args)
        json_data = get_json_data(os.path.abspath(parsed_args[JSON_ARG]))
        doc = make_docxtemplate(os.path.abspath(parsed_args[TEMPLATE_ARG]))
        doc = render_docx(doc, json_data)
        save_file(doc, parsed_args)
    except RuntimeError as e:
        print(f"Ошибка: {e}")
        return
    finally:
        if not parsed_args[QUIET_ARG]:
            print("Завершение работы программы!")


if __name__ == "__main__":
    main()

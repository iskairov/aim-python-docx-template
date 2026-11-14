# -*- coding: utf-8 -*-
"""
Скрипт установки пакета docxtpl с помощью setuptools.
"""
from __future__ import annotations

import os
import re

from setuptools import setup


def read(*names: str) -> dict[str, str]:
    """
    Читает содержимое текстовых файлов документации (.rst) в кодировке UTF-8.

    :param names: Имена файлов без расширения .rst.

    :return: Словарь {имя_файла: содержимое}.
    """

    values: dict[str, str] = {}
    for name in names:
        filename = name + ".rst"
        if os.path.isfile(filename):
            with open(filename, encoding="utf-8") as fd:
                value = fd.read()
        else:
            value = ""
        values[name] = value
    return values


long_description = """
%(README)s

История изменений
=================
%(CHANGES)s
""" % read(
    "README", "CHANGES"
)


def get_version(pkg: str) -> str:
    """
    Извлекает строку версии __version__ из файла __init__.py указанного пакета.

    :param pkg: Имя директории пакета.

    :return: Строка версии пакета.
    """

    path = os.path.join(os.path.dirname(__file__), pkg, "__init__.py")
    with open(path, encoding="utf-8") as fh:
        m = re.search(r'^__version__\s*=\s*[\'"]([^\'"]+)[\'"]', fh.read(), re.M)
    if m:
        return m.group(1)
    raise RuntimeError(f"Не удалось найти строку __version__ в файле {path}.")


setup(
    name="docxtpl",
    version=get_version("docxtpl"),
    description="Шаблонизатор документов Word (.docx) на базе Python и Jinja2",
    long_description=long_description,
    long_description_content_type="text/x-rst",
    classifiers=[
        "Intended Audience :: Developers",
        "Development Status :: 5 - Production/Stable",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Programming Language :: Python :: 3.12",
        "Programming Language :: Python :: 3.13",
        "Programming Language :: Python :: 3.14",
    ],
    keywords="jinja2 docx word template openxml",
    url="https://github.com/elapouya/python-docx-template",
    author="Eric Lapouyade",
    license="LGPL-2.1-only",
    license_files=["LICENSE.txt"],
    packages=["docxtpl"],
    python_requires=">=3.9",
    install_requires=["python-docx>=1.1.1", "jinja2", "lxml"],
    extras_require={
        "docs": ["Sphinx", "sphinxcontrib-napoleon"],
        "subdoc": ["docxcompose"],
    },
    zip_safe=False,
)

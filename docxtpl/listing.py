# -*- coding: utf-8 -*-
"""
Модуль класса Listing для вставки многострочного экранированного текста с сохранением стиля шаблона.
"""
from __future__ import annotations

from html import escape
from typing import Any


class Listing:
    r""" Класс для вставки текста с управляющими символами (\n, \a, \t, \f) и автоматическим XML-экранированием при сохранении текущего стиля шаблона. """

    def __init__(self, text: Any) -> None:
        """
        Инициализация объекта многострочного текста Listing.

        :param text: Текст или произвольный объект для преобразования в экранированную XML-строку.
        """

        # Если передана не строка и не байты, приводим к строковому представлению
        if not isinstance(text, (str, bytes)):
            text = "" if text is None else str(text)
        if isinstance(text, bytes):
            text = text.decode("utf-8", errors="ignore")
        self.xml: str = escape(text)

    def __bool__(self) -> bool:
        """
        Проверяет, содержит ли объект Listing непустой текст.

        :return: True, если текст не пуст, иначе False.
        """

        return bool(self.xml)

    def __unicode__(self) -> str:
        """
        Возвращает экранированный текст (для обратной совместимости).

        :return: Экранированная строка.
        """

        return self.xml

    def __str__(self) -> str:
        """
        Возвращает экранированную строку текста.

        :return: Экранированная строка.
        """

        return self.xml

    def __html__(self) -> str:
        """
        Возвращает экранированную строку без повторного экранирования в режиме autoescape Jinja2.

        :return: Экранированная строка.
        """

        return self.xml

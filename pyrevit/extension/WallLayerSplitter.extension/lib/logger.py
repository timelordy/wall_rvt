# -*- coding: utf-8 -*-
"""Простой обёрточный модуль для логирования в pyRevit.

Использует стандартный pyRevit-логгер и добавляет несколько удобных
функций, которые можно вызывать из команд расширения.
"""

from pyrevit import script


def get_logger(name=None):
    """Возвращает экземпляр pyRevit-логгера.

    Если указано имя, то возвращается дочерний логгер с добавленным
    префиксом. Это помогает группировать сообщения разных команд.
    """

    base_logger = script.get_logger()
    if name:
        return base_logger.getChild(name)
    return base_logger


def log_debug(message):
    get_logger().debug(message)


def log_info(message):
    get_logger().info(message)


def log_warning(message):
    get_logger().warning(message)


def log_error(message):
    get_logger().error(message)

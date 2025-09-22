# -*- coding: utf-8 -*-
"""Простой обёрточный модуль для логирования в pyRevit.

Использует стандартный pyRevit-логгер и добавляет несколько удобных
функций, которые можно вызывать из команд расширения.

Если по какой-то причине pyRevit недоступен (например, скрипт запускают
вне среды pyRevit) или стандартная папка логов ещё не создана, модуль
включает резервный режим: создаёт структуру каталогов и пишет сообщения
в собственный файл ``WallLayerSplitter.log``. Благодаря этому журнал
появится даже у пользователей, у которых штатная папка ``%APPDATA%\\pyRevit\\Logs``
пока отсутствует.
"""

import logging
import logging.handlers
import os
import threading

try:  # typing может отсутствовать в окружении IronPython 2.x
    from typing import Optional
except Exception:  # noqa: BLE001
    Optional = None  # type: ignore[assignment]

try:  # pragma: no cover - pyRevit доступен только внутри Revit
    from pyrevit import script  # type: ignore
except Exception:  # noqa: BLE001
    script = None


_FALLBACK_LOGGER_NAME = "WallLayerSplitter"
_FALLBACK_LOGGER_LOCK = threading.Lock()
_FALLBACK_LOGGER = None  # type: Optional[logging.Logger]


def _get_default_log_dir():
    """Определить каталог, где следует хранить логи.

    Сначала пытаемся получить путь из pyRevit, если он доступен. В противном
    случае формируем путь вручную: ``%APPDATA%\\pyRevit\\Logs`` для Windows и
    ``~/.pyrevit/logs`` для других платформ.
    """

    log_dir = None

    if script:
        try:
            env = script.get_pyrevit_env()
            log_dir = getattr(env, "log_dir", None)
        except Exception:  # noqa: BLE001
            log_dir = None

    if not log_dir:
        appdata = os.environ.get("APPDATA")
        if appdata:
            log_dir = os.path.join(appdata, "pyRevit", "Logs")
        else:
            log_dir = os.path.join(os.path.expanduser("~"), ".pyrevit", "logs")

    return log_dir


def _ensure_log_dir():
    """Убедиться, что каталог логов существует."""

    log_dir = _get_default_log_dir()
    if not log_dir:
        return log_dir

    try:
        os.makedirs(log_dir)
    except Exception:  # noqa: BLE001
        if not os.path.isdir(log_dir):
            raise
    return log_dir


def _configure_fallback_logger():
    """Создать и настроить резервный логгер."""

    global _FALLBACK_LOGGER

    with _FALLBACK_LOGGER_LOCK:
        if _FALLBACK_LOGGER is not None:
            return _FALLBACK_LOGGER

        log_dir = _ensure_log_dir()
        log_file = os.path.join(log_dir, "WallLayerSplitter.log")

        logger = logging.getLogger(_FALLBACK_LOGGER_NAME)
        logger.setLevel(logging.DEBUG)
        logger.propagate = False

        if not any(isinstance(handler, logging.FileHandler) for handler in logger.handlers):
            handler = logging.handlers.RotatingFileHandler(
                log_file,
                maxBytes=2 * 1024 * 1024,
                backupCount=3,
                encoding="utf-8",
            )
            handler.setFormatter(
                logging.Formatter("%(asctime)s [%(levelname)s] %(name)s: %(message)s")
            )
            logger.addHandler(handler)

        _FALLBACK_LOGGER = logger
        return logger


def _get_base_logger():
    """Получить базовый логгер pyRevit или резервный вариант."""

    if script:
        try:
            _ensure_log_dir()
            return script.get_logger()
        except Exception:  # noqa: BLE001
            pass

    return _configure_fallback_logger()


def get_logger(name=None):
    """Возвращает экземпляр логгера.

    Если указано имя, то возвращается дочерний логгер с добавленным
    префиксом. Это помогает группировать сообщения разных команд.
    """

    base_logger = _get_base_logger()
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

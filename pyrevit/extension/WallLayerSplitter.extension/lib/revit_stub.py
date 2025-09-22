# -*- coding: utf-8 -*-
"""Заглушки для Autodesk Revit API и pyRevit.

Модуль позволяет импортировать команды вне среды Autodesk Revit. Он
определяет упрощённые классы и исключения, которые используются только для
типа и не содержат рабочей логики. При обращении к методам или атрибутам
этих объектов генерируется информативное исключение, объясняющее, что
действие недоступно вне Revit.
"""

from __future__ import annotations

import enum
from typing import Any

__all__ = [
    "ArgumentException",
    "InvalidOperationException",
    "ObjectType",
    "OperationCanceledException",
    "RevitAPIUnavailableError",
    "TaskDialog",
    "XYZ",
    "AssemblyInstance",
    "BuiltInParameter",
    "CompoundStructureLayer",
    "DesignOption",
    "ElementClassFilter",
    "ElementId",
    "FamilyInstance",
    "FilteredElementCollector",
    "HostObject",
    "IntersectionResult",
    "ISelectionFilter",
    "JoinGeometryUtils",
    "LocationCurve",
    "LocationPoint",
    "Material",
    "MaterialFunctionAssignment",
    "PartUtils",
    "Phase",
    "StorageType",
    "Transaction",
    "TransactionGroup",
    "Transform",
    "Wall",
    "WallType",
    "WallUtils",
    "WorksetId",
    "revit",
    "script",
]


class RevitAPIUnavailableError(RuntimeError):
    """Исключение для обращения к Revit API вне среды Revit."""

    def __init__(self, name):
        message = (
            "Объект Revit API '{0}' недоступен вне Autodesk Revit. "
            "Выполните команду внутри Revit через pyRevit.".format(name)
        )
        super(RevitAPIUnavailableError, self).__init__(message)


class _StubBase(object):
    """Базовый класс для простых заглушек Revit API."""

    __slots__ = ()

    def __init__(self, *args, **kwargs):
        # Экземпляры таких классов не несут функциональности, но их можно
        # создавать для нужд автотестов или статических проверок.
        super(_StubBase, self).__init__()

    def __getattr__(self, item):
        raise RevitAPIUnavailableError("{0}.{1}".format(self.__class__.__name__, item))

    def __repr__(self):
        return "<stub {0}>".format(self.__class__.__name__)


class XYZ(_StubBase):
    """Минимальная заглушка для Autodesk.Revit.DB.XYZ."""

    BasisY = None

    def Normalize(self):  # pragma: no cover - заглушка
        raise RevitAPIUnavailableError("XYZ.Normalize")


class AssemblyInstance(_StubBase):
    """Заглушка семейства на основе экземпляра сборки."""


class _DynamicAttributeMeta(type):
    """Метакласс, возвращающий имя атрибута при обращении."""

    def __getattr__(cls, item):
        return item


class BuiltInParameter(object, metaclass=_DynamicAttributeMeta):  # pragma: no cover - динамическая заглушка
    """Заглушка перечисления BuiltInParameter."""


class CompoundStructureLayer(_StubBase):
    """Заглушка слоя сложной конструкции."""

    Width = 0.0
    Function = None
    MaterialId = None


class DesignOption(_StubBase):
    """Заглушка варианта проектирования."""


class ElementClassFilter(_StubBase):
    """Заглушка фильтра по типу элемента."""


class ElementId(_StubBase):
    """Заглушка идентификатора элемента."""

    InvalidElementId = object()

    def __init__(self, value=None):
        super(ElementId, self).__init__()
        self.IntegerValue = value if isinstance(value, int) else 0

    def __int__(self):  # pragma: no cover - заглушка
        return int(self.IntegerValue)


class FamilyInstance(_StubBase):
    """Заглушка экземпляра семейств."""


class FilteredElementCollector(_StubBase):
    """Заглушка коллектора элементов."""

    def __iter__(self):  # pragma: no cover - заглушка
        return iter(())

    def OfClass(self, _):  # pragma: no cover - заглушка
        return self

    def ToElements(self):  # pragma: no cover - заглушка
        return []


class HostObject(_StubBase):
    """Заглушка базового хост-элемента."""


class IntersectionResult(_StubBase):
    """Заглушка результата пересечения."""


class JoinGeometryUtils(_StubBase):
    """Заглушка для JoinGeometryUtils."""

    @staticmethod
    def GetJoinedElements(*args, **kwargs):  # pragma: no cover - заглушка
        raise RevitAPIUnavailableError("JoinGeometryUtils.GetJoinedElements")

    @staticmethod
    def AreElementsJoined(*args, **kwargs):  # pragma: no cover - заглушка
        raise RevitAPIUnavailableError("JoinGeometryUtils.AreElementsJoined")

    @staticmethod
    def UnjoinGeometry(*args, **kwargs):  # pragma: no cover - заглушка
        raise RevitAPIUnavailableError("JoinGeometryUtils.UnjoinGeometry")


class LocationCurve(_StubBase):
    """Заглушка LocationCurve."""


class LocationPoint(_StubBase):
    """Заглушка LocationPoint."""


class Material(_StubBase):
    """Заглушка материала."""


class MaterialFunctionAssignment(object, metaclass=_DynamicAttributeMeta):  # pragma: no cover - динамическая заглушка
    """Заглушка перечисления MaterialFunctionAssignment."""


class PartUtils(_StubBase):
    """Заглушка PartUtils."""

    @staticmethod
    def IsElementAssociatedWithParts(*args, **kwargs):  # pragma: no cover - заглушка
        raise RevitAPIUnavailableError("PartUtils.IsElementAssociatedWithParts")


class Phase(_StubBase):
    """Заглушка фазы."""


class StorageType(object, metaclass=_DynamicAttributeMeta):  # pragma: no cover - динамическая заглушка
    """Заглушка перечисления StorageType."""


class Transaction(_StubBase):
    """Заглушка транзакции."""

    def Start(self, *args, **kwargs):  # pragma: no cover - заглушка
        raise RevitAPIUnavailableError("Transaction.Start")

    def Commit(self, *args, **kwargs):  # pragma: no cover - заглушка
        raise RevitAPIUnavailableError("Transaction.Commit")

    def RollBack(self, *args, **kwargs):  # pragma: no cover - заглушка
        raise RevitAPIUnavailableError("Transaction.RollBack")

    def HasStarted(self):  # pragma: no cover - заглушка
        return False


class TransactionGroup(_StubBase):
    """Заглушка группы транзакций."""

    def Start(self, *args, **kwargs):  # pragma: no cover - заглушка
        raise RevitAPIUnavailableError("TransactionGroup.Start")

    def Assimilate(self, *args, **kwargs):  # pragma: no cover - заглушка
        raise RevitAPIUnavailableError("TransactionGroup.Assimilate")

    def RollBack(self, *args, **kwargs):  # pragma: no cover - заглушка
        raise RevitAPIUnavailableError("TransactionGroup.RollBack")


class Transform(_StubBase):
    """Заглушка трансформаций."""

    @staticmethod
    def CreateTranslation(*args, **kwargs):  # pragma: no cover - заглушка
        raise RevitAPIUnavailableError("Transform.CreateTranslation")


class Wall(_StubBase):
    """Заглушка стены."""


class WallType(_StubBase):
    """Заглушка типа стены."""


class WallUtils(_StubBase):
    """Заглушка WallUtils."""

    @staticmethod
    def IsWallJoinAllowedAtEnd(*args, **kwargs):  # pragma: no cover - заглушка
        raise RevitAPIUnavailableError("WallUtils.IsWallJoinAllowedAtEnd")

    @staticmethod
    def DisallowWallJoinAtEnd(*args, **kwargs):  # pragma: no cover - заглушка
        raise RevitAPIUnavailableError("WallUtils.DisallowWallJoinAtEnd")


class WorksetId(_StubBase):
    """Заглушка WorksetId."""

    InvalidWorksetId = object()


class ArgumentException(Exception):
    """Заглушка Autodesk.Revit.Exceptions.ArgumentException."""


class InvalidOperationException(Exception):
    """Заглушка Autodesk.Revit.Exceptions.InvalidOperationException."""


class OperationCanceledException(Exception):
    """Заглушка Autodesk.Revit.Exceptions.OperationCanceledException."""


class ISelectionFilter(_StubBase):
    """Заглушка интерфейса фильтра выбора."""

    def AllowElement(self, element):  # pragma: no cover - заглушка
        raise RevitAPIUnavailableError("ISelectionFilter.AllowElement")

    def AllowReference(self, reference, position):  # pragma: no cover - заглушка
        raise RevitAPIUnavailableError("ISelectionFilter.AllowReference")


class ObjectType(enum.Enum):
    """Упрощённое перечисление ObjectType."""

    Element = 0


class TaskDialog(object):
    """Заглушка диалогового окна Revit."""

    @staticmethod
    def Show(title, message):  # pragma: no cover - заглушка
        print(u"[{0}] {1}".format(title, message))


class _OutputStub(object):
    """Минимальная заглушка вывода pyRevit."""

    def print_md(self, message, *args, **kwargs):  # pragma: no cover - заглушка
        print(message)

    def close(self):  # pragma: no cover - заглушка
        return None


class _ScriptStub(object):
    """Заглушка модуля pyrevit.script."""

    def get_output(self):  # pragma: no cover - заглушка
        return _OutputStub()

    def get_pyrevit_env(self):  # pragma: no cover - заглушка
        class _Env(object):
            log_dir = ""

        return _Env()


class _RevitStub(object):
    """Заглушка модуля pyrevit.revit."""

    uidoc = None


script = _ScriptStub()
revit = _RevitStub()

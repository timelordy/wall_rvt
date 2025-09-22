# -*- coding: utf-8 -*-
"""Разделение многослойных стен на отдельные стены по слоям."""


import enum
import math
import os
import re
import sys

LIB_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", "..", "lib"))
if LIB_DIR not in sys.path:
    sys.path.append(LIB_DIR)

REVIT_API_AVAILABLE = True
try:  # pragma: no cover - импорт работает только внутри Revit
    from Autodesk.Revit.DB import (
        XYZ,
        AssemblyInstance,
        BuiltInParameter,
        CompoundStructureLayer,
        DesignOption,
        ElementClassFilter,
        ElementId,
        FamilyInstance,
        FilteredElementCollector,
        HostObject,
        IntersectionResult,
        JoinGeometryUtils,
        LocationCurve,
        LocationPoint,
        Material,
        MaterialFunctionAssignment,
        PartUtils,
        Phase,
        StorageType,
        Transaction,
        TransactionGroup,
        Transform,
        Wall,
        WallType,
        WallUtils,
        WorksetId,
    )
    from Autodesk.Revit.Exceptions import (  # noqa: F401
        ArgumentException,
        InvalidOperationException,
        OperationCanceledException,
    )
    from Autodesk.Revit.UI import TaskDialog
    from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
except Exception:  # pragma: no cover - использование заглушек вне Revit
    from revit_stub import (  # type: ignore
        ArgumentException,
        AssemblyInstance,
        BuiltInParameter,
        CompoundStructureLayer,
        DesignOption,
        ElementClassFilter,
        ElementId,
        FamilyInstance,
        FilteredElementCollector,
        HostObject,
        ISelectionFilter,
        IntersectionResult,
        JoinGeometryUtils,
        LocationCurve,
        LocationPoint,
        Material,
        MaterialFunctionAssignment,
        ObjectType,
        OperationCanceledException,
        PartUtils,
        Phase,
        StorageType,
        TaskDialog,
        Transaction,
        TransactionGroup,
        Transform,
        Wall,
        WallType,
        WallUtils,
        WorksetId,
        XYZ,
        InvalidOperationException,
    )
    REVIT_API_AVAILABLE = False

try:  # pragma: no cover - модуль доступен только в среде pyRevit
    from pyrevit import revit, script  # type: ignore
except Exception:  # pragma: no cover - fallback на заглушки
    from revit_stub import revit, script  # type: ignore

from logger import get_logger  # noqa: E402

NOT_IN_REVIT_MESSAGE = (
    "Команда доступна только в Autodesk Revit. Запустите её через pyRevit."
)

LOGGER = get_logger("WallLayerSplitter")
OUTPUT = script.get_output()

MAX_LAYER_TYPE_NAME_LENGTH = 200
MAX_LAYER_TYPE_NAME_ATTEMPTS = 50
SAFE_LAYER_TYPE_BASE_NAME_LENGTH = 60


def make_valid_wall_type_name(raw_name, max_length=SAFE_LAYER_TYPE_BASE_NAME_LENGTH, fallback_name="Тип слоя"):
    """Очистить и безопасно обрезать имя типа стены."""

    invalid_chars = {":", ";", "{", "}", "[", "]", "|", "\\", "/", "<", ">", "?", "*", '"'}
    trimmed = (raw_name or "").strip()
    if not trimmed:
        trimmed = fallback_name

    sanitized_chars = []
    for char in trimmed:
        code = ord(char)
        if code < 32 or char in invalid_chars:
            sanitized_chars.append("_")
        else:
            sanitized_chars.append(char)

    sanitized = "".join(sanitized_chars)
    sanitized = re.sub(r"\s+", " ", sanitized).strip()

    if not sanitized:
        sanitized = fallback_name

    if max_length and len(sanitized) > max_length:
        sanitized = sanitized[:max_length].rstrip()

    if not sanitized:
        sanitized = fallback_name

    return sanitized


def get_element_id_value(element_id):
    if isinstance(element_id, ElementId):
        try:
            return element_id.IntegerValue
        except Exception:  # noqa: BLE001
            try:
                return int(element_id)
            except Exception:  # noqa: BLE001
                return None
    if isinstance(element_id, int):
        return element_id
    try:
        return int(element_id)
    except (TypeError, ValueError):
        return None


def format_element_id(element_id, fallback_value=None):
    if fallback_value is not None:
        try:
            return str(int(fallback_value))
        except (TypeError, ValueError):
            return str(fallback_value)

    value = get_element_id_value(element_id)
    if value is not None:
        return str(value)

    if isinstance(element_id, ElementId):
        try:
            return str(element_id.IntegerValue)
        except Exception:  # noqa: BLE001
            pass

    if element_id is None:
        return "неизвестно"

    return str(element_id)


def try_get_wall_type(document, wall):
    if document is None or wall is None:
        return None

    try:
        type_id = wall.GetTypeId()
        if not type_id:
            return None
        return document.GetElement(type_id)
    except Exception:  # noqa: BLE001
        return None


def try_is_element_associated_with_parts(document, element_id):
    """Безопасно проверить привязку элемента к частям (Parts)."""

    if document is None or element_id is None:
        return False, "не удалось получить документ или идентификатор элемента"

    method = getattr(PartUtils, "IsElementAssociatedWithParts", None)
    if method is None:
        LOGGER.debug(
            "PartUtils.IsElementAssociatedWithParts недоступен в текущей версии API."
        )
        return False, "API не поддерживает проверку разбивки на части"

    try:
        return bool(method(document, element_id)), ""
    except (InvalidOperationException, ArgumentException):
        return False, "API отклонило проверку разбивки на части"
    except AttributeError:
        LOGGER.debug(
            "PartUtils.IsElementAssociatedWithParts отсутствует у типа PartUtils."
        )
        return False, "API не содержит метод проверки разбивки на части"
    except Exception as error:  # noqa: BLE001
        LOGGER.debug(
            "Не удалось выполнить PartUtils.IsElementAssociatedWithParts для элемента %s: %s",
            element_id,
            error,
        )
        return False, "ошибка при обращении к API Parts"


class WallLocationReference(enum.IntEnum):
    WALL_CENTERLINE = 0
    CORE_CENTERLINE = 1
    FINISH_FACE_EXTERIOR = 2
    FINISH_FACE_INTERIOR = 3
    CORE_FACE_EXTERIOR = 4
    CORE_FACE_INTERIOR = 5


class LayerWallInfo(object):
    def __init__(self, wall, layer, index, center_offset):
        self.wall = wall
        self.layer = layer
        self.index = index
        self.center_offset = center_offset
        self._half_width = layer.Width / 2.0

    def contains_offset(self, offset, tolerance):
        return (self.center_offset - self._half_width - tolerance) <= offset <= (
            self.center_offset + self._half_width + tolerance
        )


class WallSplitResult(object):
    def __init__(
        self,
        original_wall_id,
        created_walls,
        rehosted_instances,
        unmatched_instances,
        failed_detach_instances,
        original_wall_id_value=None,
    ):
        self.original_wall_id = original_wall_id
        self.original_wall_id_value = (
            original_wall_id_value
            if original_wall_id_value is not None
            else get_element_id_value(original_wall_id)
        )
        self.created_wall_ids = created_walls or []
        self.rehosted_instance_ids = rehosted_instances or []
        self.unmatched_instance_ids = unmatched_instances or []
        self.failed_detach_instance_ids = failed_detach_instances or []


class WallSelectionFilter(ISelectionFilter):
    def AllowElement(self, element):  # noqa: N802
        return isinstance(element, Wall) and has_multiple_layers(element)

    def AllowReference(self, reference, position):  # noqa: N802
        return False


def has_multiple_layers(wall):
    if wall is None:
        return False

    document = getattr(wall, "Document", None)
    wall_type = try_get_wall_type(document, wall)
    if wall_type is None:
        return False

    try:
        structure = wall_type.GetCompoundStructure() if wall_type else None
    except Exception:  # noqa: BLE001
        return False
    return bool(structure and structure.LayerCount > 1)


class WallLayerSplitterCommand(object):
    """Портированная логика команды по разделению стен."""

    def __init__(self, document):
        self.doc = document
        self.layer_type_cache = {}
        self.skip_messages = []
        self.diagnostic_log = []
        self.wall_type_name_map = None

    def execute(self):
        if not REVIT_API_AVAILABLE:
            LOGGER.error("Команда WallLayerSplitter запущена вне Autodesk Revit.")
            TaskDialog.Show("Разделение слоев стен", NOT_IN_REVIT_MESSAGE)
            if OUTPUT is not None:
                try:
                    OUTPUT.print_md("**Ошибка:** {0}".format(NOT_IN_REVIT_MESSAGE))
                except Exception:  # pragma: no cover - резервный вывод
                    pass
            return

        ui_doc = getattr(revit, "uidoc", None)
        if ui_doc is None:
            TaskDialog.Show("Разделение слоев стен", "Не удалось получить активный документ.")
            return

        LOGGER.info("Запуск команды WallLayerSplitter.")
        self.skip_messages = []
        self.diagnostic_log = []
        self.log_diagnostic("Команда запущена.")

        target_walls = list(self.collect_target_walls(ui_doc))
        if not target_walls:
            message = "Выберите хотя бы одну стену с несколькими слоями."
            TaskDialog.Show("Разделение слоев стен", message)
            LOGGER.warning("Команда завершена: подходящих стен не найдено.")
            return

        LOGGER.info("Найдено %s стен(ы) для обработки.", len(target_walls))
        self.log_diagnostic("Найдено для обработки стен: {0}.".format(len(target_walls)))

        split_results = []

        tgroup = TransactionGroup(self.doc, "Разделение слоев стен")
        try:
            tgroup.Start()
            LOGGER.debug("TransactionGroup начата.")

            transaction = Transaction(self.doc, "Создание стен по слоям")
            try:
                transaction.Start()
                LOGGER.debug("Transaction начата.")

                for wall in target_walls:
                    wall_id = wall.Id.IntegerValue
                    base_type = try_get_wall_type(self.doc, wall)
                    if base_type is None:
                        LOGGER.info(
                            "Стена %s пропущена: не удалось получить тип стены.", wall_id
                        )
                        self.report_skip_reason(
                            wall.Id, "не удалось определить тип стены (ошибка доступа к типу)"
                        )
                        continue

                    wall_type_name = getattr(base_type, "Name", None) or "<без типа>"
                    LOGGER.info("Обработка стены %s (тип '%s').", wall_id, wall_type_name)
                    self.log_diagnostic(
                        "Стена {0}: начало обработки (тип \"{1}\").".format(wall_id, wall_type_name)
                    )
                    if not self.is_wall_type_accepted(base_type):
                        LOGGER.info("Стена %s пропущена фильтром типов.", wall_id)
                        continue

                    result = self.split_wall(wall)
                    if result:
                        split_results.append(result)
                        LOGGER.info("Стена %s успешно разделена.", wall_id)
                        self.log_diagnostic(
                            "Стена {0}: разделение завершено, создано стен {1}.".format(
                                wall_id, len(result.created_wall_ids)
                            )
                        )
                    else:
                        LOGGER.info("Стена %s не была разделена.", wall_id)
                        self.log_diagnostic("Стена {0}: разделение не выполнено.".format(wall_id))

                if not split_results:
                    skip_details = self.build_skip_details(self.skip_messages)
                    message = "Ни одна из выбранных стен не подошла для разделения."
                    if skip_details:
                        message = "{}\n\n{}".format(message, skip_details)
                    LOGGER.warning("Ни одна стена не обработана. Откат транзакций.")
                    self.log_diagnostic("Ни одна стена не была разделена, выполняется откат изменений.")
                    TaskDialog.Show("Разделение слоев стен", message)
                    transaction.RollBack()
                    tgroup.RollBack()
                    return

                transaction.Commit()
                LOGGER.debug("Transaction зафиксирована.")
            except Exception:  # noqa: BLE001
                LOGGER.exception("Ошибка во время транзакции. Выполняется откат.")
                if transaction.HasStarted():
                    transaction.RollBack()
                raise
            finally:
                if transaction.HasStarted():
                    transaction.RollBack()

            tgroup.Assimilate()
            LOGGER.debug("TransactionGroup зафиксирована.")
        except Exception as error:  # noqa: BLE001
            LOGGER.exception("Команда завершилась с ошибкой: %s", error)
            tgroup.RollBack()
            diagnostics = self.get_recent_diagnostics()
            if diagnostics:
                LOGGER.error("Диагностические сообщения перед ошибкой:\n%s", "\n".join(diagnostics))
            failure_message = self.build_error_dialog_message(error, diagnostics)
            TaskDialog.Show("Разделение слоев стен", failure_message)
            self.print_diagnostics_to_output(error, diagnostics)
            return

        self.show_summary(split_results, self.skip_messages)
        LOGGER.info("Команда WallLayerSplitter завершена успешно.")

    def collect_target_walls(self, ui_doc):
        selection = ui_doc.Selection
        document = ui_doc.Document
        selected_ids = list(selection.GetElementIds())
        walls = []

        if selected_ids:
            self.log_diagnostic(
                "Выбрано заранее элементов: {0}. Фильтрация стен.".format(len(selected_ids))
            )
            for element_id in selected_ids:
                element = document.GetElement(element_id)
                if isinstance(element, Wall) and has_multiple_layers(element):
                    walls.append(element)

        if walls:
            self.log_diagnostic("Использована предварительная выборка стен: {0}.".format(len(walls)))
            return walls

        try:
            picked_refs = selection.PickObjects(ObjectType.Element, WallSelectionFilter(), "Выберите стены для разделения")
        except OperationCanceledException:
            LOGGER.info("Выбор объектов отменён пользователем.")
            self.log_diagnostic("Пользователь отменил выбор стен.")
            return []

        for reference in picked_refs:
            element = document.GetElement(reference.ElementId)
            if isinstance(element, Wall) and has_multiple_layers(element):
                walls.append(element)

        if walls:
            self.log_diagnostic("Выбрано стен после запроса: {0}.".format(len(walls)))
        return walls

    def is_wall_type_accepted(self, wall_type):
        return wall_type is not None

    def split_wall(self, wall):
        if wall is None:
            LOGGER.debug("SplitWall: передана пустая стена.")
            return None

        original_wall_id = wall.Id
        wall_id_value = get_element_id_value(original_wall_id)
        wall_label = format_element_id(original_wall_id, wall_id_value)
        self.log_diagnostic("Стена {0}: сбор исходных данных.".format(wall_label))

        base_type = try_get_wall_type(self.doc, wall)
        if base_type is None:
            LOGGER.debug(
                "SplitWall: стена %s пропущена — не удалось получить тип стены.",
                wall_label,
            )
            self.report_skip_reason(
                original_wall_id,
                "не удалось определить тип стены (ошибка доступа к типу)",
            )
            self.log_diagnostic(
                "Стена {0}: пропущена — не удалось получить тип стены.".format(
                    wall_label
                )
            )
            return None

        structure = base_type.GetCompoundStructure() if base_type else None
        if not structure or structure.LayerCount <= 1:
            LOGGER.debug(
                "SplitWall: стена %s пропущена — состав конструкции отсутствует или один слой.",
                wall_label,
            )
            self.log_diagnostic(
                "Стена {0}: пропущена — нет составной конструкции или один слой.".format(wall_label)
            )
            return None

        location = wall.Location
        if not isinstance(location, LocationCurve) or location.Curve is None:
            LOGGER.debug("SplitWall: стена %s не имеет LocationCurve.", wall_label)
            self.log_diagnostic(
                "Стена {0}: отсутствует геометрия LocationCurve.".format(wall_label)
            )
            return None

        hosted_family_ids = wall.GetDependentElements(ElementClassFilter(FamilyInstance))
        detached_instances, failed_to_detach = self.detach_hosted_family_instances(wall, hosted_family_ids)

        detached_id_set = {instance.Id.IntegerValue for instance in detached_instances if instance}
        can_delete, delete_reason = self.can_delete_original_wall(wall, detached_id_set)
        if not can_delete:
            if failed_to_detach:
                failed_list = ", ".join(str(e.IntegerValue) for e in failed_to_detach)
                if delete_reason:
                    delete_reason = "{}; не удалось временно отвязать семейства: {}".format(delete_reason, failed_list)
                else:
                    delete_reason = "не удалось временно отвязать семейства: {}".format(failed_list)
            self.restore_family_instances_to_host(detached_instances, wall)
            self.report_skip_reason(
                original_wall_id,
                "невозможно удалить исходную стену: {}".format(delete_reason),
            )
            self.log_diagnostic(
                "Стена {0}: невозможно удалить исходную стену ({1}).".format(
                    wall_label,
                    delete_reason or "неизвестная причина",
                )
            )
            return None

        base_curve = location.Curve
        orientation = wall.Orientation
        if orientation is None or orientation.GetLength() < 1e-9:
            orientation = XYZ.BasisY
        else:
            orientation = orientation.Normalize()

        base_level_id = wall.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT).AsElementId()
        base_offset = wall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET).AsDouble()
        top_constraint_id = wall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE).AsElementId()
        top_offset = wall.get_Parameter(BuiltInParameter.WALL_TOP_OFFSET).AsDouble()
        unconnected_height = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM).AsDouble()
        is_structural = wall.get_Parameter(BuiltInParameter.WALL_STRUCTURAL_SIGNIFICANT).AsInteger() == 1
        location_line = wall.get_Parameter(BuiltInParameter.WALL_KEY_REF_PARAM).AsInteger()

        layers = structure.GetLayers()
        wall_location_line = self.resolve_wall_location_line(location_line)
        total_thickness = self.calculate_total_thickness(layers)
        exterior_face_offset = total_thickness / 2.0
        reference_offset = self.calculate_reference_offset(structure, layers, wall_location_line, exterior_face_offset)

        created_walls = []
        layer_infos = []

        for index, layer in enumerate(layers):
            if layer.Width <= 0:
                continue

            layer_type = self.get_or_create_layer_type(base_type, layer, index)
            self.log_diagnostic(
                "Стена {0}: обработка слоя {1}, толщина {2:.3f}.".format(
                    wall_label, index + 1, layer.Width
                )
            )
            layer_center_offset = self.calculate_layer_center_offset(layers, index, exterior_face_offset)
            layer_offset_from_reference = reference_offset + layer_center_offset
            offset_curve = self.create_offset_curve(base_curve, orientation, layer_offset_from_reference)

            new_wall = self.create_wall_from_layer(
                offset_curve,
                layer_type,
                base_level_id,
                base_offset,
                top_constraint_id,
                top_offset,
                unconnected_height,
                wall.Flipped,
                is_structural,
                location_line,
            )

            self.copy_instance_parameters(wall, new_wall)
            created_walls.append(new_wall.Id)
            layer_infos.append(LayerWallInfo(new_wall, layer, index, layer_offset_from_reference))

        if not created_walls:
            TaskDialog.Show(
                "Разделение слоев стен",
                "Не удалось создать новые стены для исходной стены {}.".format(wall_label),
            )
            self.restore_family_instances_to_host(detached_instances, wall)
            self.log_diagnostic("Стена {0}: новые стены не созданы.".format(wall_label))
            return None

        try:
            self.log_diagnostic("Стена {0}: удаление исходной стены.".format(wall_label))
            self.doc.Delete(original_wall_id)
        except InvalidOperationException as ex:
            self.restore_family_instances_to_host(detached_instances, wall)
            raise InvalidOperationException(
                "Не удалось удалить исходную стену (ID: {}). Подробности: {}".format(
                    wall_label,
                    ex.Message,
                )
            )
        except ArgumentException as ex:
            self.restore_family_instances_to_host(detached_instances, wall)
            raise InvalidOperationException(
                "Не удалось удалить исходную стену (ID: {}). Подробности: {}".format(
                    wall_label,
                    ex.Message,
                )
            )
        except Exception as ex:  # noqa: BLE001
            self.restore_family_instances_to_host(detached_instances, wall)
            raise InvalidOperationException(
                "Не удалось удалить исходную стену (ID: {}). Подробности: {}".format(
                    wall_label,
                    ex,
                )
            )

        rehosted_instances, unmatched_instances = self.rehost_family_instances(
            base_curve,
            orientation,
            layer_infos,
            detached_instances,
        )

        result = WallSplitResult(
            original_wall_id,
            created_walls,
            rehosted_instances,
            unmatched_instances,
            [element_id for element_id in failed_to_detach] if failed_to_detach else [],
            wall_id_value,
        )

        return result

    def get_or_create_layer_type(self, base_type, layer, index):
        base_type_name = base_type.Name if isinstance(base_type, WallType) else "<без типа>"
        self.log_diagnostic(
            "  Слой {0}: поиск или создание типа на основе \"{1}\".".format(
                index + 1, base_type_name
            )
        )
        cache_key = self.build_layer_type_key(base_type.Id, layer, index)
        cached_id = self.layer_type_cache.get(cache_key)
        if cached_id:
            existing = self.doc.GetElement(cached_id)
            if isinstance(existing, WallType):
                self.log_diagnostic(
                    "  Слой {0}: найден кешированный тип (ID {1}).".format(
                        index + 1, existing.Id.IntegerValue
                    )
                )
                return existing
            self.layer_type_cache.pop(cache_key, None)

        type_name = self.build_layer_type_name(base_type, layer, index)
        for candidate in FilteredElementCollector(self.doc).OfClass(WallType).ToElements():
            if isinstance(candidate, WallType) and candidate.Name.lower() == type_name.lower():
                self.layer_type_cache[cache_key] = candidate.Id
                self.log_diagnostic(
                    "  Слой {0}: найден существующий тип \"{1}\".".format(index + 1, candidate.Name)
                )
                return candidate

        existing = self.find_wall_type_by_name(type_name)
        if isinstance(existing, WallType):
            self.layer_type_cache[cache_key] = existing.Id
            self.log_diagnostic(
                "  Слой {0}: использован тип \"{1}\" из словаря.".format(index + 1, existing.Name)
            )
            return existing

        duplicated = self.duplicate_wall_type_with_safe_name(base_type, type_name, index)
        if not isinstance(duplicated, WallType):
            raise InvalidOperationException("Не удалось продублировать тип стены для слоя.")

        new_structure = duplicated.GetCompoundStructure()
        new_layers = [CompoundStructureLayer(layer.Width, layer.Function, layer.MaterialId)]
        new_structure.SetLayers(new_layers)
        duplicated.SetCompoundStructure(new_structure)
        self.set_structural_material(duplicated, layer.MaterialId)
        self.log_diagnostic(
            "  Слой {0}: создан новый тип \"{1}\" (ID {2}).".format(
                index + 1, duplicated.Name, duplicated.Id.IntegerValue
            )
        )

        self.layer_type_cache[cache_key] = duplicated.Id
        return duplicated

    def duplicate_wall_type_with_safe_name(self, base_type, desired_name, layer_index):
        base_name = self.prepare_layer_type_base_name(desired_name)
        name_map = self.get_wall_type_name_map()

        for attempt in range(1, MAX_LAYER_TYPE_NAME_ATTEMPTS + 1):
            raw_candidate_name = self.build_candidate_layer_type_name(base_name, attempt)
            candidate_name = make_valid_wall_type_name(
                raw_candidate_name,
                max_length=MAX_LAYER_TYPE_NAME_LENGTH,
            )
            name_key = self.normalize_wall_type_name(candidate_name)
            if name_key in name_map:
                LOGGER.debug(
                    "Тип стены с именем '%s' (исходно '%s') уже существует, попытка %s пропущена.",
                    candidate_name,
                    raw_candidate_name,
                    attempt,
                )
                self.log_diagnostic(
                    "  Слой {0}: имя \"{1}\" → \"{2}\" уже занято (попытка {3}).".format(
                        layer_index + 1,
                        raw_candidate_name,
                        candidate_name,
                        attempt,
                    )
                )
                continue

            try:
                self.log_diagnostic(
                    "  Слой {0}: попытка {1} создать тип \"{2}\" (исходное \"{3}\").".format(
                        layer_index + 1,
                        attempt,
                        candidate_name,
                        raw_candidate_name,
                    )
                )
                duplicated = base_type.Duplicate(candidate_name)

            except (ArgumentException, InvalidOperationException) as error:
                LOGGER.warning(
                    "Не удалось создать тип стены '%s' (исходно '%s', попытка %s): %s",
                    candidate_name,
                    raw_candidate_name,
                    attempt,
                    error,
                )
                self.log_diagnostic(
                    "  Слой {0}: ошибка при создании типа \"{1}\" (из \"{2}\", попытка {3}) — {4}.".format(
                        layer_index + 1,
                        candidate_name,
                        raw_candidate_name,
                        attempt,
                        self.extract_error_message(error),
                    )
                )
                continue

            if isinstance(duplicated, WallType):
                self.register_wall_type(duplicated)
                self.log_diagnostic(
                    "  Слой {0}: успешно создан тип \"{1}\" (ID {2}).".format(
                        layer_index + 1, duplicated.Name, duplicated.Id.IntegerValue
                    )
                )
                return duplicated

            raise InvalidOperationException(
                "Получен некорректный тип при дублировании слоя {}.".format(layer_index + 1)
            )

        message = (
            "Не удалось подобрать уникальное имя типа для слоя {}. Сократите имена материалов или исходных типов.".format(
                layer_index + 1
            )
        )
        self.log_diagnostic("  Слой {0}: все попытки исчерпаны без успеха.".format(layer_index + 1))
        raise InvalidOperationException(message)

    def prepare_layer_type_base_name(self, desired_name):
        raw_name = (desired_name or "").strip()
        base_name = make_valid_wall_type_name(
            raw_name,
            max_length=SAFE_LAYER_TYPE_BASE_NAME_LENGTH,
        )
        if base_name != raw_name:
            LOGGER.debug(
                "Базовое имя типа слоя скорректировано: '%s' → '%s'",
                raw_name,
                base_name,
            )
        return base_name

    def build_candidate_layer_type_name(self, base_name, attempt):
        if attempt <= 1:
            return base_name
        suffix = " ({})".format(attempt)
        max_base_length = max(1, MAX_LAYER_TYPE_NAME_LENGTH - len(suffix))
        trimmed = base_name[:max_base_length].rstrip()
        if not trimmed:
            trimmed = base_name[:max_base_length]
        return "{}{}".format(trimmed, suffix)

    def get_wall_type_name_map(self):
        if self.wall_type_name_map is None:
            self.wall_type_name_map = {}
            for candidate in FilteredElementCollector(self.doc).OfClass(WallType):
                if isinstance(candidate, WallType) and candidate.Name:
                    name_key = self.normalize_wall_type_name(candidate.Name)
                    if name_key:
                        self.wall_type_name_map[name_key] = candidate
        return self.wall_type_name_map

    def register_wall_type(self, wall_type):
        if not isinstance(wall_type, WallType) or not wall_type.Name:
            return
        name_key = self.normalize_wall_type_name(wall_type.Name)
        if name_key:
            self.get_wall_type_name_map()[name_key] = wall_type

    def find_wall_type_by_name(self, type_name):
        if not type_name:
            return None
        return self.get_wall_type_name_map().get(self.normalize_wall_type_name(type_name))

    @staticmethod
    def normalize_wall_type_name(name):
        return (name or "").strip().lower()

    @staticmethod
    def calculate_total_thickness(layers):
        return sum(layer.Width for layer in layers)

    @staticmethod
    def calculate_layer_center_offset(layers, layer_index, exterior_face_offset):
        cumulative = sum(layers[i].Width for i in range(layer_index))
        center_from_exterior = cumulative + layers[layer_index].Width / 2.0
        return exterior_face_offset - center_from_exterior

    def calculate_reference_offset(self, structure, layers, wall_location_line, exterior_face_offset):
        if wall_location_line == WallLocationReference.WALL_CENTERLINE:
            return 0
        if wall_location_line == WallLocationReference.FINISH_FACE_EXTERIOR:
            return exterior_face_offset
        if wall_location_line == WallLocationReference.FINISH_FACE_INTERIOR:
            return -exterior_face_offset
        if wall_location_line == WallLocationReference.CORE_FACE_EXTERIOR:
            return self.calculate_core_face_exterior_offset(structure, layers, exterior_face_offset)
        if wall_location_line == WallLocationReference.CORE_FACE_INTERIOR:
            return self.calculate_core_face_interior_offset(structure, layers, exterior_face_offset)
        if wall_location_line == WallLocationReference.CORE_CENTERLINE:
            return self.calculate_core_centerline_offset(structure, layers, exterior_face_offset)
        return 0

    @staticmethod
    def calculate_core_face_exterior_offset(structure, layers, exterior_face_offset):
        success, exterior_thickness, _, _ = try_get_core_thicknesses(structure, layers)
        if not success:
            return 0
        return exterior_face_offset - exterior_thickness

    @staticmethod
    def calculate_core_face_interior_offset(structure, layers, exterior_face_offset):
        success, _, _, interior_thickness = try_get_core_thicknesses(structure, layers)
        if not success:
            return 0
        return -exterior_face_offset + interior_thickness

    @staticmethod
    def calculate_core_centerline_offset(structure, layers, exterior_face_offset):
        success, exterior_thickness, core_thickness, _ = try_get_core_thicknesses(structure, layers)
        if not success:
            return 0
        return exterior_face_offset - (exterior_thickness + core_thickness / 2.0)

    @staticmethod
    def resolve_wall_location_line(parameter_value):
        try:
            return WallLocationReference(parameter_value)
        except ValueError:
            return WallLocationReference.WALL_CENTERLINE

    def create_wall_from_layer(
        self,
        curve,
        wall_type,
        base_level_id,
        base_offset,
        top_constraint_id,
        top_offset,
        unconnected_height,
        flipped,
        structural,
        location_line,
    ):
        new_wall = Wall.Create(self.doc, curve, wall_type.Id, base_level_id, unconnected_height, base_offset, flipped, structural)

        top_constraint_param = new_wall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE)
        if top_constraint_id != ElementId.InvalidElementId:
            try_set_parameter(top_constraint_param, lambda: top_constraint_param.Set(top_constraint_id))
        else:
            unconnected_height_param = new_wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM)
            try_set_parameter(unconnected_height_param, lambda: unconnected_height_param.Set(unconnected_height))

        top_offset_param = new_wall.get_Parameter(BuiltInParameter.WALL_TOP_OFFSET)
        try_set_parameter(top_offset_param, lambda: top_offset_param.Set(top_offset))

        base_offset_param = new_wall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET)
        try_set_parameter(base_offset_param, lambda: base_offset_param.Set(base_offset))

        base_constraint_param = new_wall.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT)
        try_set_parameter(base_constraint_param, lambda: base_constraint_param.Set(base_level_id))

        location_line_param = new_wall.get_Parameter(BuiltInParameter.WALL_KEY_REF_PARAM)
        try_set_parameter(location_line_param, lambda: location_line_param.Set(location_line))

        return new_wall

    @staticmethod
    def create_offset_curve(base_curve, wall_orientation, offset):
        if math.fabs(offset) < 1e-9:
            return base_curve
        translation = wall_orientation.Multiply(offset)
        transform = Transform.CreateTranslation(translation)
        return base_curve.CreateTransformed(transform)

    def copy_instance_parameters(self, source_wall, target_wall):
        for built_in_parameter in DEFAULT_PARAMETERS_TO_COPY:
            source_param = source_wall.get_Parameter(built_in_parameter)
            target_param = target_wall.get_Parameter(built_in_parameter)
            if source_param is None or target_param is None:
                continue
            if not source_param.HasValue or target_param.IsReadOnly:
                continue

            storage_type = source_param.StorageType
            if storage_type == StorageType.Integer:
                value = source_param.AsInteger()
                try_set_parameter(target_param, lambda value=value: target_param.Set(value))
            elif storage_type == StorageType.Double:
                value = source_param.AsDouble()
                try_set_parameter(target_param, lambda value=value: target_param.Set(value))
            elif storage_type == StorageType.String:
                value = source_param.AsString()
                try_set_parameter(target_param, lambda value=value: target_param.Set(value))
            elif storage_type == StorageType.ElementId:
                value = source_param.AsElementId()
                try_set_parameter(target_param, lambda value=value: target_param.Set(value))

    def detach_hosted_family_instances(self, wall, hosted_family_ids):
        detached_instances = []
        failed_to_detach = []

        if not hosted_family_ids:
            return detached_instances, failed_to_detach

        for hosted_id in hosted_family_ids:
            family_instance = self.doc.GetElement(hosted_id)
            if not isinstance(family_instance, FamilyInstance):
                continue
            if not is_hosted_by_wall(family_instance, wall.Id):
                continue

            host_parameter = family_instance.get_Parameter(BuiltInParameter.HOST_ID_PARAM)
            if host_parameter is None or host_parameter.IsReadOnly:
                failed_to_detach.append(hosted_id)
                continue

            try:
                if host_parameter.Set(ElementId.InvalidElementId):
                    detached_instances.append(family_instance)
                else:
                    failed_to_detach.append(hosted_id)
            except (InvalidOperationException, ArgumentException):
                failed_to_detach.append(hosted_id)

        return detached_instances, failed_to_detach

    @staticmethod
    def restore_family_instances_to_host(family_instances, host_wall):
        if not family_instances or host_wall is None:
            return

        for family_instance in family_instances:
            if family_instance is None or not family_instance.IsValidObject:
                continue
            try_rehost_family_instance(family_instance, host_wall)

    def rehost_family_instances(self, base_curve, wall_orientation, layer_infos, detached_instances):
        rehosted_ids = []
        unmatched_ids = []

        if not layer_infos or not detached_instances:
            return rehosted_ids, unmatched_ids

        for family_instance in detached_instances:
            if family_instance is None or not family_instance.IsValidObject:
                continue

            target_layer = self.select_layer_for_instance(family_instance, base_curve, wall_orientation, layer_infos)
            if target_layer is None:
                unmatched_ids.append(family_instance.Id)
                continue

            if try_rehost_family_instance(family_instance, target_layer.wall):
                rehosted_ids.append(family_instance.Id)
            else:
                unmatched_ids.append(family_instance.Id)

        return rehosted_ids, unmatched_ids

    def select_layer_for_instance(self, family_instance, base_curve, wall_orientation, layer_infos):
        tolerance = 1e-6
        offset = try_get_instance_offset(family_instance, base_curve, wall_orientation)
        if offset is not None:
            candidates = [info for info in layer_infos if info.contains_offset(offset, tolerance)]
            if candidates:
                candidates.sort(key=lambda info: abs(info.center_offset - offset))
                return candidates[0]

        return choose_layer_by_function(layer_infos)

    def can_delete_original_wall(self, wall, detached_family_ids):
        reason = ""
        if self.doc is None or wall is None:
            reason = "не удалось получить данные стены."
            self.log_diagnostic("Блокирующая проверка: не удалось получить данные стены.")
            return False, reason

        detected_reasons = []
        detached_joined_ids = set()
        blocked_joined_ids = set()
        join_detach_failures = []

        if self.doc.IsModifiable:
            detached_joined = self.try_detach_joined_elements(wall, join_detach_failures, blocked_joined_ids)
            for element_id in detached_joined:
                detached_joined_ids.add(element_id.IntegerValue)
        else:
            self.log_diagnostic("Документ не в режиме редактирования, разрыв соединений невозможен.")

        if join_detach_failures:
            failure_list = ", ".join(sorted(set(join_detach_failures)))
            description = "не удалось разорвать соединение со следующими элементами: " + failure_list
            self.add_blocking_reason(detected_reasons, description)

        if not self.doc.IsModifiable:
            description = "документ открыт только для чтения — изменения недоступны"
            self.add_blocking_reason(detected_reasons, description)

        if self.doc.IsWorkshared:
            owner_name = get_element_owner_name(wall)
            current_user = get_document_username(self.doc)
            if is_owned_by_another_user(owner_name, current_user):
                if owner_name:
                    description = "стена занята пользователем \"{0}\"".format(owner_name)
                else:
                    description = "стена занята другим пользователем"
                self.add_blocking_reason(detected_reasons, description)
                self.log_diagnostic(
                    "Стена {0}: {1}.".format(
                        format_element_id(wall.Id),
                        description,
                    )
                )

            workset_id = wall.WorksetId
            if workset_id != WorksetId.InvalidWorksetId:
                workset_table = self.doc.GetWorksetTable()
                workset = workset_table.GetWorkset(workset_id)
                if workset is not None and not workset.IsEditable:
                    description = "рабочий набор \"{}\" не передан вам".format(workset.Name)
                    self.add_blocking_reason(detected_reasons, description)

        if wall.Pinned:
            self.add_blocking_reason(detected_reasons, "стена закреплена командой Pin")

        if wall.GroupId != ElementId.InvalidElementId:
            self.add_blocking_reason(detected_reasons, "стена входит в группу")

        design_option_id = get_design_option_id(wall)
        try:
            active_option = self.doc.ActiveDesignOptionId
        except (AttributeError, InvalidOperationException):
            # Свойство ActiveDesignOptionId появилось в API Revit сравнительно недавно,
            # поэтому на старых версиях приложения его может не быть. В этом случае,
            # а также если Revit запрещает читать активную дизайн-опцию, считаем,
            # что активной опции нет (ElementId.InvalidElementId).
            active_option = ElementId.InvalidElementId
        if design_option_id != ElementId.InvalidElementId and design_option_id != active_option:
            option_description = build_design_option_description(self.doc, design_option_id)
            description = "стена принадлежит неактивной дизайн-опции {}".format(option_description)
            self.add_blocking_reason(detected_reasons, description)

        assembly_id = wall.AssemblyInstanceId
        if assembly_id and assembly_id != ElementId.InvalidElementId:
            assembly_description = build_assembly_description(self.doc, assembly_id)
            self.add_blocking_reason(detected_reasons, "стена входит в сборку {}".format(assembly_description))

        has_associated_parts, parts_check_message = try_is_element_associated_with_parts(
            self.doc, wall.Id
        )
        if parts_check_message:
            self.log_diagnostic(
                "Проверка разбивки на части: {0}.".format(parts_check_message)
            )
        if has_associated_parts:
            self.add_blocking_reason(detected_reasons, "стена разбита на части (Parts)")

        phase_created_param = wall.get_Parameter(BuiltInParameter.WALL_PHASE_CREATED)
        if phase_created_param and phase_created_param.HasValue:
            phase_created_id = phase_created_param.AsElementId()
            active_phase_id = self.doc.ActiveView.PhaseId if self.doc.ActiveView else ElementId.InvalidElementId
            if active_phase_id != ElementId.InvalidElementId and active_phase_id != phase_created_id:
                phase_description = build_phase_description(self.doc, phase_created_id)
                description = "стена создана в фазе {}, отличной от фазы активного вида".format(phase_description)
                self.add_blocking_reason(detected_reasons, description)

        if not detected_reasons:
            dependent_elements = wall.GetDependentElements(None)
            if dependent_elements:
                blocking_descriptions = []
                for dependent_id in dependent_elements:
                    if dependent_id is None or dependent_id == ElementId.InvalidElementId:
                        continue
                    if dependent_id == wall.Id:
                        continue
                    if detached_family_ids and dependent_id.IntegerValue in detached_family_ids:
                        continue
                    if dependent_id.IntegerValue in detached_joined_ids:
                        continue
                    element = self.doc.GetElement(dependent_id)
                    if element is None or not element.IsValidObject:
                        continue
                    blocking_descriptions.append(build_element_description(element))

                if blocking_descriptions:
                    description = "у стены остались зависимые элементы: " + ", ".join(sorted(set(blocking_descriptions)))
                    self.add_blocking_reason(detected_reasons, description)

        if detected_reasons:
            reason = "; ".join(detected_reasons)
            return False, reason

        return True, ""

    def try_detach_joined_elements(self, wall, failure_messages, blocked_element_ids):
        detached_elements = []
        joined_element_ids = JoinGeometryUtils.GetJoinedElements(self.doc, wall)
        if not joined_element_ids:
            return detached_elements

        for joined_id in joined_element_ids:
            if joined_id is None or joined_id == ElementId.InvalidElementId:
                continue

            joined_element = self.doc.GetElement(joined_id)
            if not can_element_be_safely_unjoined(joined_element):
                if joined_id.IntegerValue not in blocked_element_ids:
                    failure_messages.append(build_element_description(joined_element))
                blocked_element_ids.add(joined_id.IntegerValue)
                continue

            try:
                if try_unjoin_geometry(self.doc, wall, joined_element):
                    detached_elements.append(joined_id)
                    if isinstance(joined_element, Wall):
                        try_disallow_wall_joins_at_both_ends(wall)
                        try_disallow_wall_joins_at_both_ends(joined_element)
                else:
                    failure_messages.append(build_element_description(joined_element))
            except Exception:  # noqa: BLE001
                failure_messages.append(build_element_description(joined_element))

        return detached_elements

    def show_summary(self, results, skipped_messages):
        builder = ["Результат разделения слоев стен:"]
        total_created = 0
        total_rehosted = 0
        total_unmatched = 0
        total_failed_detach = 0

        for result in results:
            total_created += len(result.created_wall_ids)
            total_rehosted += len(result.rehosted_instance_ids)
            total_unmatched += len(result.unmatched_instance_ids)
            total_failed_detach += len(result.failed_detach_instance_ids)

            builder.append(
                "Стена {0} -> создано {1} стен.".format(
                    format_element_id(result.original_wall_id, result.original_wall_id_value),
                    len(result.created_wall_ids),
                )
            )
            if result.rehosted_instance_ids:
                builder.append("    Перепривязано семейств: {0}.".format(len(result.rehosted_instance_ids)))
            if result.unmatched_instance_ids:
                unmatched_list = ", ".join(
                    format_element_id(element_id) for element_id in result.unmatched_instance_ids
                )
                builder.append("    Не удалось перепривязать автоматически: {0}.".format(unmatched_list))
            if result.failed_detach_instance_ids:
                failed_list = ", ".join(
                    format_element_id(element_id) for element_id in result.failed_detach_instance_ids
                )
                builder.append("    Не удалось временно отвязать: {0}.".format(failed_list))

        if skipped_messages:
            builder.append("")
            builder.append("Стены, которые не удалось обработать:")
            for message in sorted(set(skipped_messages)):
                builder.append("- {0}".format(message))

        summary_text = "\n".join(builder)
        TaskDialog.Show("Разделение слоев стен", summary_text)
        OUTPUT.print_md("``{}``".format(summary_text.replace("\n", "\n")))

    def report_skip_reason(self, wall_id, reason):
        if wall_id is None or not reason:
            return
        formatted = format_skip_reason(reason)
        wall_label = format_element_id(wall_id)
        message = "Стена {0}: {1}".format(wall_label, formatted)
        if message not in self.skip_messages:
            self.skip_messages.append(message)
        self.log_diagnostic(
            "Пропуск стены {0}: {1}".format(
                wall_label,
                formatted.replace("\n", " "),
            )
        )

    def build_skip_details(self, skipped_messages):
        if not skipped_messages:
            return ""
        lines = ["Стены, которые не удалось обработать:"]
        for message in sorted(set(skipped_messages)):
            lines.append("- {0}".format(message))
        return "\n".join(lines)

    def add_blocking_reason(self, detected_reasons, reason):
        if reason:
            normalized = reason.strip()
            if normalized and normalized not in detected_reasons:
                detected_reasons.append(normalized)
                self.log_diagnostic("Блокирующая проверка: {}".format(normalized))

    def log_diagnostic(self, message):
        if not message:
            return
        self.diagnostic_log.append(message)
        if len(self.diagnostic_log) > 100:
            self.diagnostic_log = self.diagnostic_log[-100:]

    def get_recent_diagnostics(self, limit=10):
        if not self.diagnostic_log:
            return []
        if not limit or limit <= 0:
            return list(self.diagnostic_log)
        return self.diagnostic_log[-limit:]

    @staticmethod
    def extract_error_message(error):
        if error is None:
            return ""
        message = getattr(error, "Message", None)
        if message:
            return str(message)
        return str(error)

    @staticmethod
    def extract_error_type_name(error):
        if error is None:
            return ""
        try:
            get_type = getattr(error, "GetType", None)
            if callable(get_type):
                dotnet_type = get_type()
                if dotnet_type is not None:
                    return str(dotnet_type.FullName)
        except Exception:  # noqa: BLE001
            pass
        return error.__class__.__name__

    def build_error_dialog_message(self, error, diagnostics):
        error_message = self.extract_error_message(error) or "Неизвестная ошибка"
        error_type = self.extract_error_type_name(error)

        lines = [
            "Произошла ошибка при разделении стен.",
        ]
        if error_type:
            lines.append("Тип: {0}".format(error_type))
        if error_message:
            lines.append("Сообщение: {0}".format(error_message))

        if diagnostics:
            lines.append("")
            lines.append("Последние действия:")
            for entry in diagnostics:
                lines.append("- {0}".format(entry))

        lines.append("")
        lines.append("Подробности см. в журнале pyRevit.")
        return "\n".join(lines)

    def print_diagnostics_to_output(self, error, diagnostics):
        if error is not None:
            error_message = self.extract_error_message(error).replace("`", "'")
            error_type = self.extract_error_type_name(error)
            OUTPUT.print_md("**Ошибка:** `{0}` ({1})".format(error_message, error_type))

        if diagnostics:
            formatted = "\n".join("* {0}".format(entry) for entry in diagnostics)
            OUTPUT.print_md("### Диагностика\n{0}".format(formatted))
        else:
            OUTPUT.print_md("Диагностическая информация отсутствует. Проверьте журнал pyRevit.")

    def build_layer_type_key(self, base_type_id, layer, index):
        material_id = layer.MaterialId.IntegerValue if layer.MaterialId else -1
        return "{0}-{1}-{2}-{3}".format(base_type_id.IntegerValue, material_id, index, round(layer.Width, 8))

    def set_structural_material(self, wall_type, material_id):
        if material_id == ElementId.InvalidElementId:
            return
        structural_param = wall_type.get_Parameter(BuiltInParameter.STRUCTURAL_MATERIAL_PARAM)
        try_set_parameter(structural_param, lambda: structural_param.Set(material_id))

    def build_layer_type_name(self, base_type, layer, index):
        material_name = "Без материала"
        if layer.MaterialId != ElementId.InvalidElementId:
            material = base_type.Document.GetElement(layer.MaterialId)
            if isinstance(material, Material) and material.Name:
                material_name = material.Name
        millimeters_per_foot = 304.8
        width_mm = layer.Width * millimeters_per_foot
        components = [
            sanitize_name_component(base_type.Name),
            "Слой {0}".format(index + 1),
            sanitize_name_component(str(layer.Function)),
            sanitize_name_component(material_name),
            "{0:.0f}мм".format(width_mm),
        ]
        raw_name = " - ".join([part for part in components if part])
        return make_valid_wall_type_name(raw_name)


DEFAULT_PARAMETERS_TO_COPY = []


def build_default_parameters():
    parameters = [
        BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS,
        BuiltInParameter.ALL_MODEL_MARK,
        BuiltInParameter.WALL_BASE_OFFSET,
        BuiltInParameter.WALL_BASE_CONSTRAINT,
        BuiltInParameter.WALL_TOP_OFFSET,
        BuiltInParameter.WALL_HEIGHT_TYPE,
        BuiltInParameter.WALL_USER_HEIGHT_PARAM,
        BuiltInParameter.WALL_STRUCTURAL_SIGNIFICANT,
        BuiltInParameter.WALL_ATTR_ROOM_BOUNDING,
    ]
    for name in ("WALL_FIRE_RATING_PARAM", "WALL_ATTR_FIRE_RATING"):
        try:
            param = getattr(BuiltInParameter, name)
            if param not in parameters:
                parameters.append(param)
        except AttributeError:
            continue
    return parameters


DEFAULT_PARAMETERS_TO_COPY = build_default_parameters()


def try_get_core_thicknesses(structure, layers):
    if structure is None:
        return False, 0.0, 0.0, 0.0
    exterior = 0.0
    core = 0.0
    interior = 0.0
    first_core = structure.GetFirstCoreLayerIndex()
    last_core = structure.GetLastCoreLayerIndex()
    if first_core < 0 or last_core < 0 or last_core < first_core:
        return False, 0.0, 0.0, 0.0

    for index, layer in enumerate(layers):
        width = layer.Width
        if index < first_core:
            exterior += width
        elif index > last_core:
            interior += width
        else:
            core += width

    return True, exterior, core, interior


def try_get_instance_offset(family_instance, reference_curve, wall_orientation):
    if family_instance is None or reference_curve is None or wall_orientation is None:
        return None

    location = family_instance.Location
    if isinstance(location, LocationPoint):
        return project_offset(location.Point, reference_curve, wall_orientation)
    if isinstance(location, LocationCurve) and location.Curve is not None:
        midpoint = location.Curve.Evaluate(0.5, True)
        return project_offset(midpoint, reference_curve, wall_orientation)
    return None


def project_offset(point, reference_curve, wall_orientation):
    if point is None:
        return None
    projection = reference_curve.Project(point)
    if not isinstance(projection, IntersectionResult):
        return None
    difference = point - projection.XYZPoint
    return difference.DotProduct(wall_orientation)


def choose_layer_by_function(layer_infos):
    if not layer_infos:
        return None
    for info in layer_infos:
        if info.layer.Function == MaterialFunctionAssignment.Structure:
            return info
    return max(layer_infos, key=lambda info: info.layer.Width)


def try_rehost_family_instance(family_instance, new_host):
    if family_instance is None or new_host is None:
        return False
    host_param = family_instance.get_Parameter(BuiltInParameter.HOST_ID_PARAM)
    if host_param is None or host_param.IsReadOnly:
        return False
    try:
        return host_param.Set(new_host.Id)
    except (InvalidOperationException, ArgumentException):
        return False


def try_set_parameter(parameter, setter):
    if parameter is None or parameter.IsReadOnly:
        return
    try:
        setter()
    except (InvalidOperationException, ArgumentException):
        pass


def is_hosted_by_wall(family_instance, wall_id):
    if family_instance is None or wall_id is None:
        return False
    host = family_instance.Host
    if host is not None and host.Id == wall_id:
        return True
    host_face = family_instance.HostFace
    if host_face is not None and host_face.ElementId == wall_id:
        return True
    return False


def can_element_be_safely_unjoined(element):
    if element is None:
        return False
    return isinstance(element, HostObject)


def try_unjoin_geometry(document, first, second):
    if document is None or first is None or second is None:
        return False
    try:
        if not JoinGeometryUtils.AreElementsJoined(document, first, second):
            return True
    except (InvalidOperationException, ArgumentException):
        return True

    try:
        JoinGeometryUtils.UnjoinGeometry(document, first, second)
        try:
            return not JoinGeometryUtils.AreElementsJoined(document, first, second)
        except (InvalidOperationException, ArgumentException):
            return True
    except (InvalidOperationException, ArgumentException):
        return False


def try_disallow_wall_joins_at_both_ends(wall):
    if wall is None:
        return
    for end_index in range(2):
        try:
            if WallUtils.IsWallJoinAllowedAtEnd(wall, end_index):
                WallUtils.DisallowWallJoinAtEnd(wall, end_index)
        except (InvalidOperationException, ArgumentException):
            continue


def get_document_username(document):
    if document is None:
        return ""
    try:
        application = document.Application
    except Exception:  # noqa: BLE001
        application = None
    if application is None:
        return ""
    try:
        username = application.Username
    except Exception:  # noqa: BLE001
        username = getattr(application, "Username", None)
    if not username:
        return ""
    return str(username).strip()


def get_element_owner_name(element):
    if element is None:
        return ""
    try:
        owner_param = element.get_Parameter(BuiltInParameter.EDITED_BY)
    except Exception:  # noqa: BLE001
        owner_param = None
    if owner_param is None:
        return ""
    try:
        owner_value = owner_param.AsString()
    except Exception:  # noqa: BLE001
        owner_value = None
    if not owner_value:
        return ""
    return str(owner_value).strip()


def normalize_username(username):
    if not username:
        return ""
    return str(username).strip().lower()


def is_owned_by_another_user(owner_name, current_username):
    owner = normalize_username(owner_name)
    if not owner:
        return False
    current = normalize_username(current_username)
    if not current:
        return True
    return owner != current


def build_element_description(element):
    if element is None:
        return "неизвестный элемент"
    category_name = element.Category.Name if element.Category else ""
    element_name = element.Name or ""
    identifier = element.Id.IntegerValue
    if category_name and element_name:
        return "{0} \"{1}\" (ID {2})".format(category_name, element_name, identifier)
    if element_name:
        return "{0} (ID {1})".format(element_name, identifier)
    return "ID {0}".format(identifier)


def get_design_option_id(element):
    if element is None:
        return ElementId.InvalidElementId
    try:
        design_option = element.DesignOption
        if design_option is not None:
            return design_option.Id
    except InvalidOperationException:
        pass
    option_param = element.get_Parameter(BuiltInParameter.DESIGN_OPTION_ID)
    return option_param.AsElementId() if option_param else ElementId.InvalidElementId


def build_design_option_description(document, design_option_id):
    if document is None or design_option_id is None or design_option_id == ElementId.InvalidElementId:
        identifier = design_option_id.IntegerValue if isinstance(design_option_id, ElementId) else 0
        return "ID {0}".format(identifier)
    element = document.GetElement(design_option_id)
    if isinstance(element, DesignOption) and element.Name:
        return '\"{0}\" (ID {1})'.format(element.Name, design_option_id.IntegerValue)
    return "ID {0}".format(design_option_id.IntegerValue)


def build_assembly_description(document, assembly_id):
    if document is None or assembly_id is None or assembly_id == ElementId.InvalidElementId:
        identifier = assembly_id.IntegerValue if isinstance(assembly_id, ElementId) else 0
        return "ID {0}".format(identifier)
    element = document.GetElement(assembly_id)
    if isinstance(element, AssemblyInstance) and element.Name:
        return '\"{0}\" (ID {1})'.format(element.Name, assembly_id.IntegerValue)
    return "ID {0}".format(assembly_id.IntegerValue)


def build_phase_description(document, phase_id):
    if document is None or phase_id is None or phase_id == ElementId.InvalidElementId:
        identifier = phase_id.IntegerValue if isinstance(phase_id, ElementId) else 0
        return "ID {0}".format(identifier)
    element = document.GetElement(phase_id)
    if isinstance(element, Phase) and element.Name:
        return '\"{0}\" (ID {1})'.format(element.Name, phase_id.IntegerValue)
    return "ID {0}".format(phase_id.IntegerValue)


def sanitize_name_component(value):
    if not value:
        return ""
    invalid_chars = {':', ';', '{', '}', '[', ']', '|', '\\', '/', '<', '>', '?', '*', '"'}
    return ''.join('_' if ch in invalid_chars else ch for ch in value).strip()


def format_skip_reason(reason):
    if not reason:
        return ""
    trimmed = reason.strip().rstrip('.')
    parts = [part.strip() for part in trimmed.split(';') if part.strip()]
    if len(parts) <= 1:
        return ensure_sentence_ending(trimmed)
    return "\n".join('• ' + ensure_sentence_ending(part) for part in parts)


def ensure_sentence_ending(text):
    if not text:
        return ""
    text = text.strip()
    if not text.endswith('.'):
        text += '.'
    return text


def main():
    command = WallLayerSplitterCommand(revit.doc)
    command.execute()


if __name__ == "__main__":
    main()

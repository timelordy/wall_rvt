using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using Autodesk.Revit.DB;
using Autodesk.Revit.Exceptions;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace WallRvt.Scripts
{
    /// <summary>
    /// Команда, разбивающая выбранные многослойные стены на отдельные стены по слоям.
    /// </summary>
    /// <remarks>
    /// Логика команды разбита на небольшие методы, чтобы облегчить расширение —
    /// например, для добавления фильтрации типов стен или предварительного просмотра результата.
    /// </remarks>
    public class WallLayerSplitter : IExternalCommand
    {
        private readonly Dictionary<string, ElementId> _layerTypeCache = new Dictionary<string, ElementId>();

        /// <inheritdoc />
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDocument = commandData.Application.ActiveUIDocument;
            if (uiDocument == null)
            {
                message = "Не удалось получить активный документ.";
                return Result.Failed;
            }

            Document document = uiDocument.Document;
            try
            {
                IList<Wall> targetWalls = CollectTargetWalls(uiDocument).ToList();
                if (!targetWalls.Any())
                {
                    const string noWallsMessage = "Выберите хотя бы одну стену с несколькими слоями.";
                    TaskDialog.Show("Разделение слоев стен", noWallsMessage);
                    message = noWallsMessage;
                    return Result.Cancelled;
                }

                IList<WallSplitResult> splitResults = new List<WallSplitResult>();

                using (TransactionGroup transactionGroup = new TransactionGroup(document, "Разделение слоёв стен"))
                {
                    transactionGroup.Start();

                    try
                    {
                        using (Transaction transaction = new Transaction(document, "Создание стен по слоям"))
                        {
                            transaction.Start();

                            foreach (Wall wall in targetWalls)
                            {
                                if (!IsWallTypeAccepted(wall.WallType))
                                {
                                    continue;
                                }

                                WallSplitResult result = SplitWall(document, wall);
                                if (result != null)
                                {
                                    splitResults.Add(result);
                                }
                            }

                            if (!splitResults.Any())
                            {
                                const string nothingProcessed = "Ни одна из выбранных стен не подошла для разделения.";
                                TaskDialog.Show("Разделение слоев стен", nothingProcessed);
                                transaction.RollBack();
                                transactionGroup.RollBack();
                                message = nothingProcessed;
                                return Result.Cancelled;
                            }

                            transaction.Commit();
                        }

                        transactionGroup.Assimilate();
                    }
                    catch (Exception ex)
                    {
                        transactionGroup.RollBack();
                        message = ex.Message;
                        TaskDialog.Show("Разделение слоев стен", $"Ошибка: {ex.Message}");
                        return Result.Failed;
                    }
                }

                ShowSummary(splitResults);
                return Result.Succeeded;
            }
            catch (OperationCanceledException)
            {
                message = "Выбор объектов отменён пользователем.";
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                TaskDialog.Show("Разделение слоев стен", $"Ошибка: {ex.Message}");
                return Result.Failed;
            }
        }

        /// <summary>
        /// Собирает список стен, подходящих для разделения.
        /// </summary>
        /// <param name="uiDocument">Активный документ Revit.</param>
        /// <returns>Перечень выбранных стен.</returns>
        /// <remarks>
        /// Метод сначала пытается использовать текущее выделение.
        /// Если оно пустое, пользователю предлагается выбрать стены вручную.
        /// </remarks>
        protected virtual IEnumerable<Wall> CollectTargetWalls(UIDocument uiDocument)
        {
            Selection selection = uiDocument.Selection;
            Document document = uiDocument.Document;
            IList<ElementId> selectedIds = selection.GetElementIds().ToList();
            IList<Wall> walls = new List<Wall>();

            if (selectedIds.Any())
            {
                foreach (ElementId id in selectedIds)
                {
                    if (document.GetElement(id) is Wall wall && HasMultipleLayers(wall))
                    {
                        walls.Add(wall);
                    }
                }
            }

            if (walls.Any())
            {
                return walls;
            }

            IList<Reference> pickedReferences = selection.PickObjects(ObjectType.Element, new WallSelectionFilter(),
                "Выберите стены для разделения");

            foreach (Reference reference in pickedReferences)
            {
                if (document.GetElement(reference.ElementId) is Wall wall && HasMultipleLayers(wall))
                {
                    walls.Add(wall);
                }
            }

            return walls;
        }

        /// <summary>
        /// Определяет, подходит ли тип стены для обработки.
        /// </summary>
        /// <param name="wallType">Тип стены.</param>
        /// <returns>Возвращает true, если тип разрешён, иначе false.</returns>
        /// <remarks>
        /// Метод оставлен виртуальным, чтобы потом можно было реализовать фильтрацию конкретных типов стен.
        /// </remarks>
        protected virtual bool IsWallTypeAccepted(WallType wallType) => wallType != null;

        private WallSplitResult SplitWall(Document document, Wall wall)
        {
            CompoundStructure structure = wall.WallType.GetCompoundStructure();
            if (structure == null || structure.LayerCount <= 1)
            {
                return null;
            }

            LocationCurve locationCurve = wall.Location as LocationCurve;
            if (locationCurve == null)
            {
                return null;
            }

            IList<ElementId> createdWalls = new List<ElementId>();
            Curve baseCurve = locationCurve.Curve;
            XYZ wallOrientation = wall.Orientation;
            double orientationLength = wallOrientation.GetLength();
            wallOrientation = orientationLength > 1e-9 ? wallOrientation.Normalize() : XYZ.BasisY;

            ElementId baseLevelId = wall.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT).AsElementId();
            double baseOffset = wall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET).AsDouble();
            ElementId topConstraintId = wall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE).AsElementId();
            double topOffset = wall.get_Parameter(BuiltInParameter.WALL_TOP_OFFSET).AsDouble();
            double unconnectedHeight = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM).AsDouble();
            bool isStructural = wall.get_Parameter(BuiltInParameter.WALL_STRUCTURAL_SIGNIFICANT).AsInteger() == 1;
            int locationLine = wall.get_Parameter(BuiltInParameter.WALL_KEY_REF_PARAM).AsInteger();

            IList<CompoundStructureLayer> layers = structure.GetLayers();
            for (int index = 0; index < layers.Count; index++)
            {
                CompoundStructureLayer layer = layers[index];
                if (layer.Width <= 0)
                {
                    continue;
                }

                WallType layerType = GetOrCreateLayerType(document, wall.WallType, layer, index);

                double offset = structure.GetOffsetForLayer(index);
                XYZ translation = wallOrientation.Multiply(offset);
                Transform transform = Transform.CreateTranslation(translation);
                Curve translatedCurve = baseCurve.CreateTransformed(transform);

                Wall newWall = CreateWallFromLayer(document, translatedCurve, layerType, baseLevelId, baseOffset,
                    topConstraintId, topOffset, unconnectedHeight, wall.Flipped, isStructural, locationLine);

                CopyInstanceParameters(wall, newWall);
                createdWalls.Add(newWall.Id);
            }

            document.Delete(wall.Id);
            return new WallSplitResult(wall.Id, createdWalls);
        }

        private static string BuildLayerTypeKey(ElementId baseTypeId, CompoundStructureLayer layer, int index)
        {
            return string.Join("_", baseTypeId.IntegerValue, index, layer.Function,
                layer.MaterialId.IntegerValue, layer.Width.ToString("F6", CultureInfo.InvariantCulture));
        }

        /// <summary>
        /// Получает уже созданный тип стены для слоя или создаёт новый.
        /// </summary>
        private WallType GetOrCreateLayerType(Document document, WallType baseType, CompoundStructureLayer layer, int index)
        {
            string cacheKey = BuildLayerTypeKey(baseType.Id, layer, index);
            if (_layerTypeCache.TryGetValue(cacheKey, out ElementId cachedTypeId))
            {
                if (document.GetElement(cachedTypeId) is WallType cachedType)
                {
                    return cachedType;
                }

                _layerTypeCache.Remove(cacheKey);
            }

            string typeName = BuildLayerTypeName(baseType, layer, index);
            WallType existingType = new FilteredElementCollector(document)
                .OfClass(typeof(WallType))
                .Cast<WallType>()
                .FirstOrDefault(t => string.Equals(t.Name, typeName, StringComparison.OrdinalIgnoreCase));

            if (existingType != null)
            {
                _layerTypeCache[cacheKey] = existingType.Id;
                return existingType;
            }

            WallType duplicatedType = (WallType)baseType.Duplicate(typeName);
            CompoundStructure newStructure = duplicatedType.GetCompoundStructure();
            IList<CompoundStructureLayer> newLayers = new List<CompoundStructureLayer>
            {
                new CompoundStructureLayer(layer.Width, layer.Function, layer.MaterialId)
            };

            newStructure.SetLayers(newLayers);
            newStructure.StructuralMaterialId = layer.MaterialId;
            duplicatedType.SetCompoundStructure(newStructure);

            _layerTypeCache[cacheKey] = duplicatedType.Id;
            return duplicatedType;
        }

        private static string BuildLayerTypeName(WallType baseType, CompoundStructureLayer layer, int index)
        {
            string materialName = "Без материала";
            if (layer.MaterialId != ElementId.InvalidElementId)
            {
                Material material = baseType.Document.GetElement(layer.MaterialId) as Material;
                if (material != null)
                {
                    materialName = material.Name;
                }
            }

            const double millimetersPerFoot = 304.8;
            double widthMillimeters = layer.Width * millimetersPerFoot;
            string function = layer.Function.ToString();
            return $"{baseType.Name} | Слой {index + 1} | {function} | {materialName} | {widthMillimeters:F0}мм";
        }

        private Wall CreateWallFromLayer(Document document, Curve curve, WallType wallType, ElementId baseLevelId,
            double baseOffset, ElementId topConstraintId, double topOffset, double unconnectedHeight, bool flipped,
            bool structural, int locationLine)
        {
            Wall newWall = Wall.Create(document, curve, wallType.Id, baseLevelId, unconnectedHeight, baseOffset, flipped, structural);

            Parameter topConstraintParam = newWall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE);
            if (topConstraintParam != null && topConstraintId != ElementId.InvalidElementId)
            {
                topConstraintParam.Set(topConstraintId);
            }
            else
            {
                Parameter unconnectedHeightParam = newWall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM);
                unconnectedHeightParam?.Set(unconnectedHeight);
            }

            Parameter topOffsetParam = newWall.get_Parameter(BuiltInParameter.WALL_TOP_OFFSET);
            topOffsetParam?.Set(topOffset);

            Parameter baseOffsetParam = newWall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET);
            baseOffsetParam?.Set(baseOffset);

            Parameter baseConstraintParam = newWall.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT);
            baseConstraintParam?.Set(baseLevelId);

            Parameter locationLineParam = newWall.get_Parameter(BuiltInParameter.WALL_KEY_REF_PARAM);
            locationLineParam?.Set(locationLine);

            return newWall;
        }

        private static void CopyInstanceParameters(Wall source, Wall target)
        {
            IList<BuiltInParameter> parametersToCopy = new List<BuiltInParameter>
            {
                BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS,
                BuiltInParameter.ALL_MODEL_MARK,
                BuiltInParameter.WALL_BASE_OFFSET,
                BuiltInParameter.WALL_TOP_OFFSET,
                BuiltInParameter.WALL_STRUCTURAL_SIGNIFICANT
            };

            foreach (BuiltInParameter builtInParameter in parametersToCopy)
            {
                Parameter sourceParameter = source.get_Parameter(builtInParameter);
                Parameter targetParameter = target.get_Parameter(builtInParameter);
                if (sourceParameter != null && targetParameter != null && !targetParameter.IsReadOnly)
                {
                    switch (sourceParameter.StorageType)
                    {
                        case StorageType.Double:
                            targetParameter.Set(sourceParameter.AsDouble());
                            break;
                        case StorageType.Integer:
                            targetParameter.Set(sourceParameter.AsInteger());
                            break;
                        case StorageType.String:
                            targetParameter.Set(sourceParameter.AsString());
                            break;
                        case StorageType.ElementId:
                            targetParameter.Set(sourceParameter.AsElementId());
                            break;
                    }
                }
            }
        }

        private static bool HasMultipleLayers(Wall wall)
        {
            CompoundStructure structure = wall.WallType.GetCompoundStructure();
            return structure != null && structure.LayerCount > 1;
        }

        private static void ShowSummary(IEnumerable<WallSplitResult> results)
        {
            StringBuilder builder = new StringBuilder();
            builder.AppendLine("Результат разделения слоёв стен:");
            int totalCreated = 0;

            foreach (WallSplitResult result in results)
            {
                int createdCount = result.CreatedWallIds.Count;
                totalCreated += createdCount;
                builder.AppendLine($"- Стена {result.OriginalWallId.IntegerValue}: создано {createdCount} стен");
            }

            builder.AppendLine($"Всего новых стен: {totalCreated}");
            TaskDialog.Show("Разделение слоев стен", builder.ToString());
        }

        private class WallSelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element elem)
            {
                return elem is Wall wall && HasMultipleLayers(wall);
            }

            public bool AllowReference(Reference reference, XYZ position) => false;
        }

        private class WallSplitResult
        {
            public WallSplitResult(ElementId originalWallId, IList<ElementId> createdWalls)
            {
                OriginalWallId = originalWallId;
                CreatedWallIds = createdWalls;
            }

            public ElementId OriginalWallId { get; }

            public IList<ElementId> CreatedWallIds { get; }
        }
    }
}

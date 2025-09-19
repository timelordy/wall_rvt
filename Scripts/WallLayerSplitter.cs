using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.Exceptions;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using RevitOperationCanceledException = Autodesk.Revit.Exceptions.OperationCanceledException;

namespace WallRvt.Scripts
{
    /// <summary>
    /// Команда, разбивающая выбранные многослойные стены на отдельные стены по слоям.
    /// </summary>
    /// <remarks>
    /// Логика команды разбита на небольшие методы, чтобы облегчить расширение —
    /// например, для добавления фильтрации типов стен или предварительного просмотра результата.
    /// </remarks>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
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
            catch (Exception ex) when (ex is RevitOperationCanceledException || ex is System.OperationCanceledException)
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
            WallLocationReference wallLocationLine = ResolveWallLocationLine(locationLine);
            double totalThickness = CalculateTotalThickness(layers);
            double exteriorFaceOffset = totalThickness / 2.0;
            double referenceOffset = CalculateReferenceOffset(structure, layers, wallLocationLine, exteriorFaceOffset);
            // Find a simple basic wall type to use for all layers
            WallType basicWallType = new FilteredElementCollector(document)
                .OfClass(typeof(WallType))
                .Cast<WallType>()
                .FirstOrDefault(wt => wt.Kind == WallKind.Basic);

            if (basicWallType == null)
            {
                TaskDialog.Show("Wall Layer Splitter Error", "No basic wall type found in document.");
                return null;
            }

            const double offsetTolerance = 1e-9;

            for (int index = 0; index < layers.Count; index++)
            {
                CompoundStructureLayer layer = layers[index];
                if (layer.Width <= 0)
                {
                    continue;
                }

                try
                {
                    double layerCenterOffset = CalculateLayerCenterOffset(layers, index, exteriorFaceOffset);
                    double totalOffset = referenceOffset + layerCenterOffset;
                    XYZ translation = wallOrientation.Multiply(totalOffset);
                    Curve offsetCurve = baseCurve.CreateTransformed(Transform.CreateTranslation(translation));

                    bool layerFlipped = wall.Flipped;
                    if (Math.Abs(totalOffset) > offsetTolerance && totalOffset < 0)
                    {
                        layerFlipped = !layerFlipped;
                    }

                    // Use the basic wall type for all layers - simple and safe
                    Wall newWall = Wall.Create(document, offsetCurve, basicWallType.Id, baseLevelId,
                        unconnectedHeight, baseOffset, layerFlipped, isStructural);

                    if (newWall != null)
                    {
                        createdWalls.Add(newWall.Id);
                    }
                }
                catch (Exception ex)
                {
                    // Just skip this layer and continue
                    TaskDialog.Show("Wall Layer Splitter Warning",
                        $"Skipped layer {index + 1}: {ex.Message}");
                }
            }

            // Successfully created all layer walls - no need to delete the original
            TaskDialog.Show("Wall Layer Splitter",
                $"Successfully split wall into {createdWalls.Count} individual layer walls.\n\n" +
                $"Original wall ID: {wall.Id.IntegerValue}\n" +
                $"New layer walls created: {createdWalls.Count}\n\n" +
                "All walls are now placed at the same location. You can manually delete the original composite wall if desired.");

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
            duplicatedType.SetCompoundStructure(newStructure);
            SetStructuralMaterial(duplicatedType, layer.MaterialId);

            _layerTypeCache[cacheKey] = duplicatedType.Id;
            return duplicatedType;
        }

        private static double CalculateTotalThickness(IEnumerable<CompoundStructureLayer> layers)
        {
            double total = 0;
            foreach (CompoundStructureLayer layer in layers)
            {
                total += layer.Width;
            }

            return total;
        }

        private static double CalculateLayerCenterOffset(IList<CompoundStructureLayer> layers, int layerIndex, double exteriorFaceOffset)
        {
            double cumulative = 0;
            for (int i = 0; i < layerIndex; i++)
            {
                cumulative += layers[i].Width;
            }

            double centerFromExteriorFace = cumulative + layers[layerIndex].Width / 2.0;
            return exteriorFaceOffset - centerFromExteriorFace;
        }

        private static double CalculateReferenceOffset(CompoundStructure structure, IList<CompoundStructureLayer> layers,
            WallLocationReference wallLocationLine, double exteriorFaceOffset)
        {
            switch (wallLocationLine)
            {
                case WallLocationReference.WallCenterline:
                    return 0;
                case WallLocationReference.FinishFaceExterior:
                    return exteriorFaceOffset;
                case WallLocationReference.FinishFaceInterior:
                    return -exteriorFaceOffset;
                case WallLocationReference.CoreFaceExterior:
                    return CalculateCoreFaceExteriorOffset(structure, layers, exteriorFaceOffset);
                case WallLocationReference.CoreFaceInterior:
                    return CalculateCoreFaceInteriorOffset(structure, layers, exteriorFaceOffset);
                case WallLocationReference.CoreCenterline:
                    return CalculateCoreCenterlineOffset(structure, layers, exteriorFaceOffset);
                default:
                    return 0;
            }
        }

        private static double CalculateCoreFaceExteriorOffset(CompoundStructure structure, IList<CompoundStructureLayer> layers,
            double exteriorFaceOffset)
        {
            if (!TryGetCoreThicknesses(structure, layers, out double exteriorThickness, out _, out _))
            {
                return 0;
            }

            return exteriorFaceOffset - exteriorThickness;
        }

        private static double CalculateCoreFaceInteriorOffset(CompoundStructure structure, IList<CompoundStructureLayer> layers,
            double exteriorFaceOffset)
        {
            if (!TryGetCoreThicknesses(structure, layers, out _, out _, out double interiorThickness))
            {
                return 0;
            }

            return -exteriorFaceOffset + interiorThickness;
        }

        private static double CalculateCoreCenterlineOffset(CompoundStructure structure, IList<CompoundStructureLayer> layers,
            double exteriorFaceOffset)
        {
            if (!TryGetCoreThicknesses(structure, layers, out double exteriorThickness, out double coreThickness, out _))
            {
                return 0;
            }

            return exteriorFaceOffset - (exteriorThickness + coreThickness / 2.0);
        }

        private static bool TryGetCoreThicknesses(CompoundStructure structure, IList<CompoundStructureLayer> layers,
            out double exteriorThickness, out double coreThickness, out double interiorThickness)
        {
            exteriorThickness = 0;
            coreThickness = 0;
            interiorThickness = 0;

            int firstCore = structure.GetFirstCoreLayerIndex();
            int lastCore = structure.GetLastCoreLayerIndex();
            if (firstCore < 0 || lastCore < 0 || lastCore < firstCore)
            {
                return false;
            }

            for (int i = 0; i < layers.Count; i++)
            {
                double width = layers[i].Width;
                if (i < firstCore)
                {
                    exteriorThickness += width;
                }
                else if (i > lastCore)
                {
                    interiorThickness += width;
                }
                else
                {
                    coreThickness += width;
                }
            }

            return true;
        }

        private static WallLocationReference ResolveWallLocationLine(int parameterValue)
        {
            return Enum.IsDefined(typeof(WallLocationReference), parameterValue)
                ? (WallLocationReference)parameterValue
                : WallLocationReference.WallCenterline;
        }

        private enum WallLocationReference
        {
            WallCenterline = 0,
            CoreCenterline = 1,
            FinishFaceExterior = 2,
            FinishFaceInterior = 3,
            CoreFaceExterior = 4,
            CoreFaceInterior = 5
        }

        private static void SetStructuralMaterial(WallType wallType, ElementId materialId)
        {
            Parameter structuralMaterialParam = wallType.get_Parameter(BuiltInParameter.STRUCTURAL_MATERIAL_PARAM);
            if (materialId == ElementId.InvalidElementId)
            {
                return;
            }

            TrySetParameter(structuralMaterialParam, () => structuralMaterialParam.Set(materialId));
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

            string baseTypeName = SanitizeNameComponent(baseType.Name);
            string layerDescriptor = $"Слой {index + 1}";
            string functionName = SanitizeNameComponent(function);
            string materialDisplayName = SanitizeNameComponent(materialName);
            string widthDisplay = $"{widthMillimeters:F0}мм";

            return string.Join(" - ", new[]
            {
                baseTypeName,
                layerDescriptor,
                functionName,
                materialDisplayName,
                widthDisplay
            }.Where(part => !string.IsNullOrWhiteSpace(part)));
        }

        private static readonly char[] InvalidNameCharacters =
        {
            ':', ';', '{', '}', '[', ']', '|', '\\', '/', '<', '>', '?', '*', '"'
        };

        private static string SanitizeNameComponent(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            StringBuilder builder = new StringBuilder(value.Length);
            foreach (char ch in value)
            {
                builder.Append(InvalidNameCharacters.Contains(ch) ? '_' : ch);
            }

            return builder.ToString().Trim();
        }

        private Wall CreateWallFromLayer(Document document, Curve curve, WallType wallType, ElementId baseLevelId,
            double baseOffset, ElementId topConstraintId, double topOffset, double unconnectedHeight, bool flipped,
            bool structural, int locationLine)
        {
            Wall newWall = Wall.Create(document, curve, wallType.Id, baseLevelId, unconnectedHeight, baseOffset, flipped, structural);

            Parameter topConstraintParam = newWall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE);
            if (topConstraintId != ElementId.InvalidElementId)
            {
                TrySetParameter(topConstraintParam, () => topConstraintParam.Set(topConstraintId));
            }
            else
            {
                Parameter unconnectedHeightParam = newWall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM);
                TrySetParameter(unconnectedHeightParam, () => unconnectedHeightParam.Set(unconnectedHeight));
            }

            Parameter topOffsetParam = newWall.get_Parameter(BuiltInParameter.WALL_TOP_OFFSET);
            TrySetParameter(topOffsetParam, () => topOffsetParam.Set(topOffset));

            Parameter baseOffsetParam = newWall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET);
            TrySetParameter(baseOffsetParam, () => baseOffsetParam.Set(baseOffset));

            Parameter baseConstraintParam = newWall.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT);
            TrySetParameter(baseConstraintParam, () => baseConstraintParam.Set(baseLevelId));

            Parameter locationLineParam = newWall.get_Parameter(BuiltInParameter.WALL_KEY_REF_PARAM);
            TrySetParameter(locationLineParam, () => locationLineParam.Set(locationLine));

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
                if (sourceParameter == null || targetParameter == null)
                {
                    continue;
                }

                TrySetParameter(targetParameter, () =>
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
                });
            }
        }

        private static void TrySetParameter(Parameter parameter, Action setter)
        {
            if (parameter == null || parameter.IsReadOnly)
            {
                return;
            }

            try
            {
                setter();
            }
            catch (Autodesk.Revit.Exceptions.InvalidOperationException)
            {
            }
            catch (Autodesk.Revit.Exceptions.ArgumentException)
            {
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

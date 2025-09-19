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
using InvalidOperationException = System.InvalidOperationException;
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
        private readonly List<string> _skipMessages = new List<string>();

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
            _skipMessages.Clear();

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

                ShowSummary(splitResults, _skipMessages);
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

            if (!CanDeleteOriginalWall(document, wall, out string deleteRestrictionReason))
            {
                ReportSkipReason(wall.Id, $"невозможно удалить исходную стену: {deleteRestrictionReason}");
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
            IList<LayerWallInfo> layerWallInfos = new List<LayerWallInfo>();

            for (int index = 0; index < layers.Count; index++)
            {
                CompoundStructureLayer layer = layers[index];
                if (layer.Width <= 0)
                {
                    continue;
                }

                try
                {

                    WallType layerType = GetOrCreateLayerType(document, wall.WallType, layer, index);
                    double layerCenterOffset = CalculateLayerCenterOffset(layers, index, exteriorFaceOffset);
                    double layerOffsetFromReference = referenceOffset + layerCenterOffset;
                    Curve offsetCurve = CreateOffsetCurve(baseCurve, wallOrientation, layerOffsetFromReference);

                    Wall newWall = CreateWallFromLayer(document, offsetCurve, layerType, baseLevelId, baseOffset,
                        topConstraintId, topOffset, unconnectedHeight, wall.Flipped, isStructural, locationLine);

                    CopyInstanceParameters(wall, newWall);
                    createdWalls.Add(newWall.Id);
                    layerWallInfos.Add(new LayerWallInfo(newWall, layer, index, layerOffsetFromReference));
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("Разделение слоев стен",
                        $"Не удалось создать стену для слоя {index + 1}: {ex.Message}");
                }
            }

            IList<ElementId> rehostedInstances;
            IList<ElementId> unmatchedInstances;

            RehostFamilyInstances(document, wall, layerWallInfos, baseCurve, wallOrientation,
                out rehostedInstances, out unmatchedInstances);

            if (!createdWalls.Any())
            {
                TaskDialog.Show("Разделение слоев стен",
                    $"Не удалось создать новые стены для исходной стены {wall.Id.IntegerValue}. Исходная стена оставлена без изменений.");
                return null;
            }

            try
            {
                document.Delete(wall.Id);
            }
            catch (Autodesk.Revit.Exceptions.InvalidOperationException ex)
            {
                string errorMessage =
                    $"Не удалось удалить исходную стену (ID: {wall.Id.IntegerValue}). Возможно, на неё ссылаются другие элементы. Подробности: {ex.Message}";
                throw new InvalidOperationException(errorMessage, ex);
            }
            catch (Autodesk.Revit.Exceptions.ArgumentException ex)
            {
                string errorMessage =
                    $"Не удалось удалить исходную стену (ID: {wall.Id.IntegerValue}). Подробности: {ex.Message}";
                throw new InvalidOperationException(errorMessage, ex);
            }
            catch (Exception ex)
            {
                string errorMessage =
                    $"Не удалось удалить исходную стену (ID: {wall.Id.IntegerValue}). Подробности: {ex.Message}";
                throw new InvalidOperationException(errorMessage, ex);
            }

            TaskDialog.Show("Разделение слоев стен",
                $"Стена {wall.Id.IntegerValue} успешно разделена на {createdWalls.Count} новых стен по слоям.\n\n" +
                "Исходная стена удалена автоматически.");

            return new WallSplitResult(wall.Id, createdWalls, rehostedInstances, unmatchedInstances);
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

        private static Curve CreateOffsetCurve(Curve baseCurve, XYZ wallOrientation, double offset)
        {
            if (Math.Abs(offset) < 1e-9)
            {
                return baseCurve;
            }

            XYZ translation = wallOrientation.Multiply(offset);
            Transform transform = Transform.CreateTranslation(translation);
            return baseCurve.CreateTransformed(transform);
        }

        private static readonly BuiltInParameter[] DefaultParametersToCopy = BuildDefaultParametersToCopy();

        private static BuiltInParameter[] BuildDefaultParametersToCopy()
        {
            List<BuiltInParameter> parameters = new List<BuiltInParameter>
            {
                BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS,
                BuiltInParameter.ALL_MODEL_MARK,
                BuiltInParameter.WALL_BASE_OFFSET,
                BuiltInParameter.WALL_BASE_CONSTRAINT,
                BuiltInParameter.WALL_TOP_OFFSET,
                BuiltInParameter.WALL_HEIGHT_TYPE,
                BuiltInParameter.WALL_USER_HEIGHT_PARAM,
                BuiltInParameter.WALL_STRUCTURAL_SIGNIFICANT,
                BuiltInParameter.WALL_ATTR_ROOM_BOUNDING,
            };

            TryAddBuiltInParameterByName(parameters, "WALL_FIRE_RATING_PARAM");
            TryAddBuiltInParameterByName(parameters, "WALL_ATTR_FIRE_RATING");

            return parameters.ToArray();
        }

        private static void TryAddBuiltInParameterByName(ICollection<BuiltInParameter> target, string parameterName)
        {
            if (Enum.TryParse(parameterName, out BuiltInParameter parsedParameter) && !target.Contains(parsedParameter))
            {
                target.Add(parsedParameter);
            }
        }

        private static void CopyInstanceParameters(Wall source, Wall target)
        {
            foreach (BuiltInParameter builtInParameter in DefaultParametersToCopy)
            {
                Parameter sourceParameter = source.get_Parameter(builtInParameter);
                Parameter targetParameter = target.get_Parameter(builtInParameter);
                CopyParameterValue(sourceParameter, targetParameter);
            }

            foreach (Parameter sourceParameter in source.Parameters.Cast<Parameter>())
            {
                if (sourceParameter == null || sourceParameter.IsReadOnly)
                {
                    continue;
                }

                Definition definition = sourceParameter.Definition;
                if (definition == null)
                {
                    continue;
                }

                Parameter targetParameter = target.get_Parameter(definition);
                if (targetParameter == null || targetParameter.IsReadOnly)
                {
                    continue;
                }

                if (sourceParameter.StorageType != targetParameter.StorageType)
                {
                    continue;
                }

                CopyParameterValue(sourceParameter, targetParameter);
            }
        }

        private static void CopyParameterValue(Parameter sourceParameter, Parameter targetParameter)
        {
            if (sourceParameter == null || targetParameter == null)
            {
                return;
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

        private void RehostFamilyInstances(Document document, Wall originalWall, IList<LayerWallInfo> layerWalls,
            Curve baseCurve, XYZ wallOrientation, out IList<ElementId> rehostedInstanceIds,
            out IList<ElementId> unmatchedInstanceIds)
        {
            rehostedInstanceIds = new List<ElementId>();
            unmatchedInstanceIds = new List<ElementId>();

            if (layerWalls == null || !layerWalls.Any())
            {
                return;
            }

#if REVIT_2024_OR_GREATER
            ICollection<ElementId> hostedElementIds = HostObjectUtils.GetDirectlyHostedElements(document, originalWall.Id);
#else
            ICollection<ElementId> hostedElementIds = GetDirectlyHostedElementsLegacy(document, originalWall);
#endif
            if (hostedElementIds == null || hostedElementIds.Count == 0)
            {
                return;
            }

            foreach (ElementId hostedId in hostedElementIds)
            {
                if (!(document.GetElement(hostedId) is FamilyInstance familyInstance))
                {
                    continue;
                }

                LayerWallInfo targetLayer = SelectLayerForInstance(familyInstance, baseCurve, wallOrientation, layerWalls);
                if (targetLayer == null)
                {
                    unmatchedInstanceIds.Add(hostedId);
                    continue;
                }

                if (TryRehostFamilyInstance(familyInstance, targetLayer.Wall))
                {
                    rehostedInstanceIds.Add(hostedId);
                }
                else
                {
                    unmatchedInstanceIds.Add(hostedId);
                }
            }
        }

#if !REVIT_2024_OR_GREATER
        private static ICollection<ElementId> GetDirectlyHostedElementsLegacy(Document document, Wall hostWall)
        {
            if (document == null || hostWall == null)
            {
                return new List<ElementId>();
            }

            IList<ElementId> directlyHosted = new FilteredElementCollector(document)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(familyInstance =>
                    (familyInstance.Host != null && familyInstance.Host.Id == hostWall.Id) ||
                    (familyInstance.HostFace != null && familyInstance.HostFace.ElementId == hostWall.Id))
                .Select(familyInstance => familyInstance.Id)
                .ToList();

            return directlyHosted;
        }
#endif

        private LayerWallInfo SelectLayerForInstance(FamilyInstance familyInstance, Curve baseCurve, XYZ wallOrientation,
            IList<LayerWallInfo> layerWalls)
        {
            const double tolerance = 1e-6;
            double? offset = TryGetInstanceOffset(familyInstance, baseCurve, wallOrientation);
            if (offset.HasValue)
            {
                LayerWallInfo matchedLayer = layerWalls
                    .Where(info => info.ContainsOffset(offset.Value, tolerance))
                    .OrderBy(info => Math.Abs(info.CenterOffset - offset.Value))
                    .FirstOrDefault();

                if (matchedLayer != null)
                {
                    return matchedLayer;
                }
            }

            return ChooseLayerByFunction(layerWalls);
        }

        private static double? TryGetInstanceOffset(FamilyInstance familyInstance, Curve referenceCurve, XYZ wallOrientation)
        {
            if (familyInstance == null || referenceCurve == null || wallOrientation == null)
            {
                return null;
            }

            if (familyInstance.Location is LocationPoint locationPoint)
            {
                return ProjectOffset(locationPoint.Point, referenceCurve, wallOrientation);
            }

            if (familyInstance.Location is LocationCurve locationCurve)
            {
                XYZ midpoint = locationCurve.Curve?.Evaluate(0.5, true);
                return ProjectOffset(midpoint, referenceCurve, wallOrientation);
            }

            return null;
        }

        private static double? ProjectOffset(XYZ point, Curve referenceCurve, XYZ wallOrientation)
        {
            if (point == null)
            {
                return null;
            }

            IntersectionResult projection = referenceCurve.Project(point);
            if (projection == null)
            {
                return null;
            }

            XYZ difference = point - projection.XYZPoint;
            return difference.DotProduct(wallOrientation);
        }

        private static LayerWallInfo ChooseLayerByFunction(IList<LayerWallInfo> layerWalls)
        {
            LayerWallInfo structuralLayer = layerWalls
                .FirstOrDefault(info => info.Layer.Function == MaterialFunctionAssignment.Structure);
            if (structuralLayer != null)
            {
                return structuralLayer;
            }

            return layerWalls
                .OrderByDescending(info => info.Layer.Width)
                .FirstOrDefault();
        }

        private static bool TryRehostFamilyInstance(FamilyInstance familyInstance, Wall newHost)
        {
            if (familyInstance == null || newHost == null)
            {
                return false;
            }

            Parameter hostParameter = familyInstance.get_Parameter(BuiltInParameter.HOST_ID_PARAM);
            if (hostParameter == null || hostParameter.IsReadOnly)
            {
                return false;
            }

            try
            {
                bool result = hostParameter.Set(newHost.Id);
                return result;
            }
            catch (Autodesk.Revit.Exceptions.InvalidOperationException)
            {
                return false;
            }
            catch (Autodesk.Revit.Exceptions.ArgumentException)
            {
                return false;
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

        private void ShowSummary(IEnumerable<WallSplitResult> results, IEnumerable<string> skippedMessages)
        {
            StringBuilder builder = new StringBuilder();
            builder.AppendLine("Результат разделения слоёв стен:");
            int totalCreated = 0;
            int totalRehosted = 0;
            int totalUnmatched = 0;

            foreach (WallSplitResult result in results)
            {
                int createdCount = result.CreatedWallIds.Count;
                int rehostedCount = result.RehostedInstanceIds.Count;
                int unmatchedCount = result.UnmatchedInstanceIds.Count;

                totalCreated += createdCount;
                totalRehosted += rehostedCount;
                totalUnmatched += unmatchedCount;

                builder.AppendLine(
                    $"- Стена {result.OriginalWallId.IntegerValue}: создано {createdCount} стен, перенесено семейств {rehostedCount}");

                if (unmatchedCount > 0)
                {
                    string unmatchedList = string.Join(", ",
                        result.UnmatchedInstanceIds.Select(id => id.IntegerValue.ToString(CultureInfo.InvariantCulture)));
                    builder.AppendLine($"  ⚠ Требуют ручной корректировки: {unmatchedList}");
                }
            }

            builder.AppendLine($"Всего новых стен: {totalCreated}");
            builder.AppendLine($"Всего перенесённых семейств: {totalRehosted}");

            if (totalUnmatched > 0)
            {
                builder.AppendLine();
                builder.AppendLine("Не все элементы удалось автоматически переназначить. Проверьте перечисленные идентификаторы.");
            }

            if (skippedMessages != null)
            {
                IList<string> skippedList = skippedMessages.Where(message => !string.IsNullOrWhiteSpace(message)).ToList();
                if (skippedList.Any())
                {
                    builder.AppendLine();
                    builder.AppendLine("Стены, которые не удалось обработать:");
                    foreach (string message in skippedList)
                    {
                        builder.AppendLine($"- {message}");
                    }
                }
            }

            TaskDialog.Show("Разделение слоев стен", builder.ToString());
        }

        private static bool CanDeleteOriginalWall(Document document, Wall wall, out string reason)
        {
            reason = string.Empty;

            if (document == null || wall == null)
            {
                reason = "не удалось получить данные стены.";
                return false;
            }

            List<string> detectedReasons = new List<string>();

            if (!document.IsModifiable)
            {
                detectedReasons.Add("документ открыт только для чтения — изменения в нём сейчас недоступны");
            }

            if (document.IsWorkshared)
            {
                CheckoutStatus checkoutStatus = WorksharingUtils.GetCheckoutStatus(document, wall.Id);
                if (checkoutStatus != CheckoutStatus.OwnedByMe && checkoutStatus != CheckoutStatus.NotOwned)
                {
                    detectedReasons.Add("стена занята другим пользователем или находится в недоступном рабочем наборе");
                }

                WorksetId worksetId = wall.WorksetId;
                if (worksetId != WorksetId.InvalidWorksetId)
                {
                    Workset workset = WorksetTable.GetWorkset(document, worksetId);
                    if (workset != null && !workset.IsEditable)
                    {
                        detectedReasons.Add($"рабочий набор \"{workset.Name}\" не передан вам для редактирования");
                    }
                }
            }

            if (wall.Pinned)
            {
                detectedReasons.Add("стена закреплена командой Pin");
            }

            if (wall.GroupId != ElementId.InvalidElementId)
            {
                detectedReasons.Add("стена входит в группу");
            }

            ICollection<ElementId> dependentElements = wall.GetDependentElements(null);
            if (dependentElements != null && dependentElements.Count > 0)
            {
                detectedReasons.Add("у стены есть зависимые элементы или ограничения (например, присоединённые перекрытия, крыши или элементы вариантов), которые блокируют удаление");
            }

            if (!detectedReasons.Any())
            {
                return true;
            }

            reason = string.Join("; ", detectedReasons) + ".";
            return false;
        }

        private void ReportSkipReason(ElementId wallId, string reason)
        {
            if (wallId == null || string.IsNullOrWhiteSpace(reason))
            {
                return;
            }

            string message = $"Стена {wallId.IntegerValue}: {reason}";
            if (!_skipMessages.Contains(message))
            {
                _skipMessages.Add(message);
            }
        }

        private class WallSelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element elem)
            {
                return elem is Wall wall && HasMultipleLayers(wall);
            }

            public bool AllowReference(Reference reference, XYZ position) => false;
        }

        private class LayerWallInfo
        {
            public LayerWallInfo(Wall wall, CompoundStructureLayer layer, int index, double centerOffset)
            {
                Wall = wall;
                Layer = layer;
                Index = index;
                CenterOffset = centerOffset;
                HalfWidth = layer.Width / 2.0;
            }

            public Wall Wall { get; }

            public CompoundStructureLayer Layer { get; }

            public int Index { get; }

            public double CenterOffset { get; }

            private double HalfWidth { get; }

            private double StartOffset => CenterOffset - HalfWidth;

            private double EndOffset => CenterOffset + HalfWidth;

            public ElementId WallId => Wall.Id;

            public bool ContainsOffset(double offset, double tolerance)
            {
                return offset >= StartOffset - tolerance && offset <= EndOffset + tolerance;
            }
        }

        private class WallSplitResult
        {
            public WallSplitResult(ElementId originalWallId, IList<ElementId> createdWalls,
                IList<ElementId> rehostedInstances, IList<ElementId> unmatchedInstances)
            {
                OriginalWallId = originalWallId;
                CreatedWallIds = createdWalls ?? new List<ElementId>();
                RehostedInstanceIds = rehostedInstances ?? new List<ElementId>();
                UnmatchedInstanceIds = unmatchedInstances ?? new List<ElementId>();
            }

            public ElementId OriginalWallId { get; }

            public IList<ElementId> CreatedWallIds { get; }

            public IList<ElementId> RehostedInstanceIds { get; }

            public IList<ElementId> UnmatchedInstanceIds { get; }
        }
    }
}

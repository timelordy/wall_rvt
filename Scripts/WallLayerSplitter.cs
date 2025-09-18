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
            for (int index = 0; index < layers.Count; index++)
            {
                CompoundStructureLayer layer = layers[index];
                if (layer.Width <= 0)
                {
                    continue;
                }

                WallType layerType = GetOrCreateLayerType(document, wall.WallType, layer, index);

                double layerCenterOffset = CalculateLayerCenterOffset(layers, index, exteriorFaceOffset);
                double offset = layerCenterOffset - referenceOffset;

                // Place all new layer walls at the exact same location as the original wall
                Curve translatedCurve = baseCurve;

                Wall newWall = CreateWallFromLayer(document, translatedCurve, layerType, baseLevelId, baseOffset,
                    topConstraintId, topOffset, unconnectedHeight, wall.Flipped, isStructural, locationLine);

                CopyInstanceParameters(wall, newWall);
                createdWalls.Add(newWall.Id);
            }

            // Handle the original wall after creating all layer walls
            bool wallDeleted = false;

            // First check if the wall can be deleted
            if (document.CanDeleteElement(wall.Id))
            {
                try
                {
                    PrepareWallForDeletion(document, wall);
                    document.Delete(wall.Id);
                    wallDeleted = true;

                    TaskDialog.Show("Wall Layer Splitter",
                        $"Successfully split wall into {createdWalls.Count} layer walls.\n\n" +
                        "The original wall has been replaced with individual layer walls at the same location.");
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("Wall Layer Splitter",
                        $"Successfully created {createdWalls.Count} layer walls, but deletion failed unexpectedly.\n\n" +
                        $"Error: {ex.Message}\n\n" +
                        "The original wall remains in place. You can manually delete it when ready.");
                }
            }
            else
            {
                // Wall cannot be deleted - offer force deletion option
                TaskDialogResult result = TaskDialog.Show("Wall Layer Splitter",
                    $"Successfully created {createdWalls.Count} layer walls.\n\n" +
                    "The original wall cannot be deleted due to dependencies or constraints.\n\n" +
                    "Would you like to try force deletion? This will break all connections and remove dependent elements.",
                    TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No);

                if (result == TaskDialogResult.Yes)
                {
                    try
                    {
                        ForceDeleteWall(document, wall);

                        // Check again if we can delete it after force preparation
                        if (document.CanDeleteElement(wall.Id))
                        {
                            document.Delete(wall.Id);
                            wallDeleted = true;
                            TaskDialog.Show("Wall Layer Splitter", "Force deletion successful! The original wall has been removed.");
                        }
                        else
                        {
                            TaskDialog.Show("Wall Layer Splitter",
                                "Force deletion preparation completed, but the wall still cannot be deleted.\n\n" +
                                "The original wall remains in place. You may need to manually resolve remaining constraints.");
                        }
                    }
                    catch (Exception forceEx)
                    {
                        TaskDialog.Show("Wall Layer Splitter",
                            $"Force deletion failed: {forceEx.Message}\n\n" +
                            "The original wall remains in place. You may need to manually delete it.");
                    }
                }
                else
                {
                    TaskDialog.Show("Wall Layer Splitter",
                        "Layer walls created successfully. The original wall remains in place.\n\n" +
                        "You can manually delete it when ready.");
                }
            }

            return new WallSplitResult(wall.Id, createdWalls);
        }

        private static bool PrepareWallForDeletion(Document document, Wall wall)
        {
            try
            {
                var diagnostics = new List<string>();

                // Step 1: Unjoin all walls connected to this wall
                var joinedWallsCount = UnjoinConnectedWalls(document, wall);
                if (joinedWallsCount > 0)
                {
                    diagnostics.Add($"Unjoined {joinedWallsCount} connected walls");
                }

                // Step 2: Handle hosted elements (doors, windows, etc.)
                var hostedElementsCount = HandleHostedElements(document, wall);
                if (hostedElementsCount > 0)
                {
                    diagnostics.Add($"Removed {hostedElementsCount} hosted elements (doors/windows)");
                }

                // Step 3: Check if wall is part of a group
                if (IsWallInGroup(wall))
                {
                    ShowDetailedError(wall, "Wall is part of a group and cannot be automatically processed", diagnostics);
                    return false;
                }

                // Step 4: Handle dependent elements (dimensions, annotations, etc.)
                var dependentElementsCount = HandleDependentElements(document, wall);
                if (dependentElementsCount > 0)
                {
                    diagnostics.Add($"Removed {dependentElementsCount} dependent elements (dimensions/annotations)");
                }

                // Step 5: Check for other potential blockers
                var remainingBlockers = DiagnoseRemainingDependencies(document, wall);
                if (remainingBlockers.Count > 0)
                {
                    ShowDetailedError(wall, "Wall has remaining dependencies", diagnostics.Concat(remainingBlockers).ToList());
                    return false;
                }

                if (diagnostics.Count > 0)
                {
                    TaskDialog.Show("Wall Layer Splitter",
                        $"Automatic processing completed for wall {wall.Id.IntegerValue}:\n" +
                        string.Join("\n", diagnostics.Select(d => "• " + d)));
                }

                return true;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Wall Layer Splitter",
                    $"Error during automatic processing of wall {wall.Id.IntegerValue}: {ex.Message}");
                return false;
            }
        }

        private static void ForceDeleteWall(Document document, Wall wall)
        {
            // Aggressive preparation for deletion - removes all constraints and dependencies
            // Note: This method does NOT delete the wall itself, just prepares it for deletion

            // 1. Unpin the wall if it's pinned (critical for Revit 2024+)
            try
            {
                if (wall.Pinned)
                {
                    wall.Pinned = false;
                }
            }
            catch { /* ignore unpinning errors */ }

            // 2. Disallow wall joins at both ends using proper loop approach
            try
            {
                // Disallow joins at start of wall (index 0)
                while (WallUtils.IsWallJoinAllowedAtEnd(wall, 0))
                {
                    WallUtils.DisallowWallJoinAtEnd(wall, 0);
                }

                // Disallow joins at end of wall (index 1)
                while (WallUtils.IsWallJoinAllowedAtEnd(wall, 1))
                {
                    WallUtils.DisallowWallJoinAtEnd(wall, 1);
                }

                // Force document regeneration to apply join changes
                document.Regenerate();
            }
            catch { /* ignore join disallowing errors */ }

            // 3. Unjoin from all possible elements more efficiently
            try
            {
                // Get only walls and structural elements that might be joined
                var potentialJoinedElements = new FilteredElementCollector(document)
                    .WhereElementIsNotElementType()
                    .Where(e => e.Id != wall.Id &&
                               (e is Wall || e.Category?.Id?.IntegerValue == (int)BuiltInCategory.OST_StructuralFraming ||
                                e.Category?.Id?.IntegerValue == (int)BuiltInCategory.OST_StructuralColumns ||
                                e.Category?.Id?.IntegerValue == (int)BuiltInCategory.OST_Floors))
                    .ToList();

                foreach (var element in potentialJoinedElements)
                {
                    try
                    {
                        if (JoinGeometryUtils.AreElementsJoined(document, wall, element))
                        {
                            JoinGeometryUtils.UnjoinGeometry(document, wall, element);
                        }
                    }
                    catch { /* ignore individual unjoin failures */ }
                }
            }
            catch { /* ignore general unjoining errors */ }

            // 4. Remove all hosted elements forcefully
            try
            {
                var hostedElements = wall.FindInserts(true, true, true, true).ToList();
                if (hostedElements.Count > 0)
                {
                    // Check which hosted elements can be deleted before attempting deletion
                    var deletableHosted = hostedElements.Where(id => document.CanDeleteElement(id)).ToList();

                    foreach (var elementId in deletableHosted)
                    {
                        try
                        {
                            // Unpin hosted element if needed
                            var hostedElement = document.GetElement(elementId);
                            if (hostedElement != null && hostedElement.Pinned)
                            {
                                hostedElement.Pinned = false;
                            }
                            document.Delete(elementId);
                        }
                        catch { /* ignore hosted element deletion failures */ }
                    }
                }
            }
            catch { /* ignore general hosted element handling */ }

            // 5. Remove all dependent elements using dependency set
            try
            {
                var dependentIds = wall.GetDependentElements(null);
                if (dependentIds.Count > 0)
                {
                    // Check which dependent elements can be deleted before attempting deletion
                    var deletableDependents = dependentIds.Where(id => document.CanDeleteElement(id)).ToList();

                    foreach (var depId in deletableDependents)
                    {
                        try
                        {
                            // Unpin dependent element if needed
                            var dependentElement = document.GetElement(depId);
                            if (dependentElement != null && dependentElement.Pinned)
                            {
                                dependentElement.Pinned = false;
                            }
                            document.Delete(depId);
                        }
                        catch { /* ignore dependent element deletion failures */ }
                    }
                }
            }
            catch { /* ignore general dependent element handling */ }

            // 6. Final regeneration before checking wall deletion eligibility
            try
            {
                document.Regenerate();
            }
            catch { /* ignore regeneration errors */ }

            // Note: Wall deletion is handled by the caller using CanDeleteElement check
        }

        private static TaskDialogResult ShowDeletionOptionsDialog(int layerCount, string errorMessage)
        {
            TaskDialog dialog = new TaskDialog("Wall Layer Splitter")
            {
                MainInstruction = $"Successfully created {layerCount} layer walls, but cannot delete original wall.",
                MainContent = $"Error: {errorMessage}\n\nChoose how to handle the original wall:",
                CommonButtons = TaskDialogCommonButtons.None
            };

            dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Force Delete",
                "Try aggressive deletion (breaks all connections and removes dependencies)");
            dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Hide Wall",
                "Hide the original wall in current view (safer option)");
            dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink3, "Keep Visible",
                "Leave the original wall in place for manual handling");

            dialog.DefaultButton = TaskDialogResult.CommandLink2;

            TaskDialogResult result = dialog.Show();

            switch (result)
            {
                case TaskDialogResult.CommandLink1:
                    return TaskDialogResult.Yes; // Force delete
                case TaskDialogResult.CommandLink2:
                    return TaskDialogResult.No; // Hide wall
                default:
                    return TaskDialogResult.Cancel; // Keep visible
            }
        }

        private static bool TryHideWall(Document document, Wall wall)
        {
            try
            {
                // Get the current view
                View currentView = document.ActiveView;
                if (currentView == null)
                {
                    // Try to get the first available view
                    var views = new FilteredElementCollector(document)
                        .OfClass(typeof(View))
                        .Cast<View>()
                        .Where(v => !v.IsTemplate && v.CanBePrinted)
                        .ToList();

                    currentView = views.FirstOrDefault();
                }

                if (currentView != null)
                {
                    // Hide the wall in the current view
                    currentView.HideElements(new List<ElementId> { wall.Id });
                    return true;
                }

                return false;
            }
            catch
            {
                return false;
            }
        }

        private static void ShowDetailedError(Wall wall, string mainMessage, List<string> diagnostics)
        {
            var message = $"{mainMessage} for wall {wall.Id.IntegerValue}";

            if (diagnostics.Count > 0)
            {
                message += "\n\nActions taken:\n" + string.Join("\n", diagnostics.Select(d => "• " + d));
            }

            message += "\n\nThis wall requires manual intervention before it can be split into layers.";

            TaskDialog.Show("Wall Layer Splitter", message);
        }

        private static int UnjoinConnectedWalls(Document document, Wall wall)
        {
            int unjoinedCount = 0;
            try
            {
                // Get all walls in the document
                var allWalls = new FilteredElementCollector(document)
                    .OfClass(typeof(Wall))
                    .Cast<Wall>()
                    .ToList();

                // Check each wall for joins with our target wall
                foreach (var otherWall in allWalls)
                {
                    if (otherWall.Id == wall.Id) continue;

                    try
                    {
                        // Check if walls are joined and unjoin them
                        if (JoinGeometryUtils.AreElementsJoined(document, wall, otherWall))
                        {
                            JoinGeometryUtils.UnjoinGeometry(document, wall, otherWall);
                            unjoinedCount++;
                        }
                    }
                    catch
                    {
                        // Ignore individual unjoin failures
                    }
                }

                // Also try to disallow joins at both ends of the wall
                try
                {
                    WallUtils.DisallowWallJoinAtEnd(wall, 0);
                    WallUtils.DisallowWallJoinAtEnd(wall, 1);
                }
                catch { }
            }
            catch
            {
                // If general unjoining fails, continue
            }
            return unjoinedCount;
        }

        private static int HandleHostedElements(Document document, Wall wall)
        {
            try
            {
                // Get all family instances that might be hosted by this wall
                var hostedElements = new FilteredElementCollector(document)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .Where(fi => fi.Host != null && fi.Host.Id == wall.Id)
                    .ToList();

                var elementsToDelete = new List<ElementId>();

                foreach (var familyInstance in hostedElements)
                {
                    // For doors and windows, we'll delete them as they can't be easily transferred
                    var category = familyInstance.Category;
                    if (category?.Id == new ElementId(BuiltInCategory.OST_Doors) ||
                        category?.Id == new ElementId(BuiltInCategory.OST_Windows))
                    {
                        elementsToDelete.Add(familyInstance.Id);
                    }
                }

                // Delete hosted elements that would prevent wall deletion
                if (elementsToDelete.Count > 0)
                {
                    document.Delete(elementsToDelete);
                    return elementsToDelete.Count;
                }

                return 0;
            }
            catch
            {
                // If hosted element handling fails, continue
                return 0;
            }
        }

        private static int HandleDependentElements(Document document, Wall wall)
        {
            int removedCount = 0;
            try
            {
                // Get all dependent elements
                var dependentElementIds = wall.GetDependentElements(null);
                var elementsToDelete = new List<ElementId>();

                foreach (var depId in dependentElementIds)
                {
                    try
                    {
                        var element = document.GetElement(depId);
                        if (element == null) continue;

                        // Check if this is a dimension, annotation, or other deletable dependent element
                        if (IsDeletableDependentElement(element))
                        {
                            elementsToDelete.Add(depId);
                        }
                    }
                    catch
                    {
                        // Skip elements that can't be accessed
                    }
                }

                // Delete the dependent elements that can be safely removed
                if (elementsToDelete.Count > 0)
                {
                    try
                    {
                        var deletedIds = document.Delete(elementsToDelete);
                        removedCount = deletedIds.Count;
                    }
                    catch
                    {
                        // Try deleting elements one by one if batch deletion fails
                        foreach (var id in elementsToDelete)
                        {
                            try
                            {
                                document.Delete(id);
                                removedCount++;
                            }
                            catch
                            {
                                // Skip elements that can't be deleted
                            }
                        }
                    }
                }
            }
            catch
            {
                // If dependent element handling fails, continue
            }

            return removedCount;
        }

        private static bool IsDeletableDependentElement(Element element)
        {
            try
            {
                var category = element.Category;
                if (category == null) return false;

                var categoryId = category.Id;

                // These categories can usually be safely deleted when splitting walls
                return categoryId == new ElementId(BuiltInCategory.OST_Dimensions) ||
                       categoryId == new ElementId(BuiltInCategory.OST_GenericAnnotation) ||
                       categoryId == new ElementId(BuiltInCategory.OST_TextNotes) ||
                       categoryId == new ElementId(BuiltInCategory.OST_DetailComponents) ||
                       categoryId == new ElementId(BuiltInCategory.OST_Tags);
            }
            catch
            {
                return false;
            }
        }

        private static bool IsWallInGroup(Wall wall)
        {
            try
            {
                return wall.GroupId != null && wall.GroupId != ElementId.InvalidElementId;
            }
            catch
            {
                return false;
            }
        }

        private static List<string> DiagnoseRemainingDependencies(Document document, Wall wall)
        {
            var blockers = new List<string>();

            try
            {
                // Check for remaining joined walls
                var allWalls = new FilteredElementCollector(document)
                    .OfClass(typeof(Wall))
                    .Cast<Wall>()
                    .ToList();

                var joinedWalls = allWalls.Where(w => w.Id != wall.Id &&
                    JoinGeometryUtils.AreElementsJoined(document, wall, w)).ToList();

                if (joinedWalls.Count > 0)
                {
                    blockers.Add($"Still joined to {joinedWalls.Count} walls: " +
                        string.Join(", ", joinedWalls.Select(w => w.Id.IntegerValue).Take(3)) +
                        (joinedWalls.Count > 3 ? "..." : ""));
                }

                // Check for remaining hosted elements
                var remainingHosted = new FilteredElementCollector(document)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .Where(fi => fi.Host != null && fi.Host.Id == wall.Id)
                    .ToList();

                if (remainingHosted.Count > 0)
                {
                    var hostCategories = remainingHosted
                        .GroupBy(h => h.Category?.Name ?? "Unknown")
                        .Select(g => $"{g.Key} ({g.Count()})")
                        .ToList();

                    blockers.Add("Remaining hosted elements: " + string.Join(", ", hostCategories));
                }

                // Check for remaining dependent elements using this wall as a reference
                var dependentElementIds = wall.GetDependentElements(null);
                if (dependentElementIds.Count > 0)
                {
                    var dependentByCategory = new Dictionary<string, int>();

                    foreach (var depId in dependentElementIds)
                    {
                        try
                        {
                            var element = document.GetElement(depId);
                            if (element?.Category != null)
                            {
                                var categoryName = element.Category.Name;
                                if (dependentByCategory.ContainsKey(categoryName))
                                    dependentByCategory[categoryName]++;
                                else
                                    dependentByCategory[categoryName] = 1;
                            }
                            else
                            {
                                if (dependentByCategory.ContainsKey("Unknown"))
                                    dependentByCategory["Unknown"]++;
                                else
                                    dependentByCategory["Unknown"] = 1;
                            }
                        }
                        catch
                        {
                            if (dependentByCategory.ContainsKey("Inaccessible"))
                                dependentByCategory["Inaccessible"]++;
                            else
                                dependentByCategory["Inaccessible"] = 1;
                        }
                    }

                    var dependentDescription = string.Join(", ",
                        dependentByCategory.Select(kvp => $"{kvp.Key} ({kvp.Value})"));

                    blockers.Add($"Still referenced by dependent elements: {dependentDescription}");
                }
            }
            catch (Exception ex)
            {
                blockers.Add($"Error during dependency check: {ex.Message}");
            }

            return blockers;
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

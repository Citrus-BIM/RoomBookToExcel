using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using Floor = Autodesk.Revit.DB.Floor;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Color = DocumentFormat.OpenXml.Spreadsheet.Color;

namespace RoomBookToExcel
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class RoomBookToExcelCommand : IExternalCommand
    {
        RoomBookToExcelProgressBarWPF roomBookToExcelProgressBarWPF;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try { _ = GetPluginStartInfo(); } catch { }

            Document doc = commandData.Application.ActiveUIDocument.Document;
            Guid roombookRoomNumber = new Guid("22868552-0e64-49b2-b8d9-9a2534bf0e14");
            Guid elemData = new Guid("659c3180-6565-41bc-a332-d82502953510");

            List<Room> roomList = new FilteredElementCollector(doc)
                .OfClass(typeof(SpatialElement))
                .WhereElementIsNotElementType()
                .Where(r => r.GetType() == typeof(Room))
                .Cast<Room>()
                .Where(r => r.Area > 0)
                .OrderBy(r => r.Number, new AlphanumComparatorFastString())
                .ToList();

            if (roomList.Count == 0)
            {
                TaskDialog.Show("Revit", "Проект не содержит помещения!");
                return Result.Cancelled;
            }

            RoomBookToExcelWPF roomBookToExcelWPF = new RoomBookToExcelWPF();
            roomBookToExcelWPF.ShowDialog();
            if (roomBookToExcelWPF.DialogResult != true) return Result.Cancelled;

            string exportOptionName = roomBookToExcelWPF.ExportOptionName;

            if (exportOptionName == "rbt_FinishingForEachRoom")
            {
                StartProgressWindow(roomList.Count);

                int step = 0;
                try
                {
                    // Диалог сохранения (как было)
                    var saveDialog = new System.Windows.Forms.SaveFileDialog();
                    saveDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
                    var dlgRes = saveDialog.ShowDialog();
                    if (dlgRes != System.Windows.Forms.DialogResult.OK)
                    {
                        roomBookToExcelProgressBarWPF.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.Close());
                        return Result.Cancelled;
                    }

                    string filePath = saveDialog.FileName;

                    // ==== OPENXML ====
                    using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
                    {
                        WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();
                        workbookPart.Workbook = new Workbook();

                        WorkbookStylesPart stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
                        stylesPart.Stylesheet = BuildStylesheet_ISOCPEUR();
                        stylesPart.Stylesheet.Save();

                        SharedStringTablePart sstPart = workbookPart.AddNewPart<SharedStringTablePart>();
                        sstPart.SharedStringTable = new SharedStringTable();

                        WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                        SheetData sheetData = new SheetData();

                        // Columns (widths)
                        Columns columns = new Columns();
                        AddColumn(columns, 1, 1, 10);
                        AddColumn(columns, 2, 2, 30);
                        AddColumn(columns, 3, 3, 10);
                        AddColumn(columns, 4, 4, 10);
                        AddColumn(columns, 5, 5, 50);
                        AddColumn(columns, 6, 6, 10);
                        AddColumn(columns, 7, 7, 15);
                        AddColumn(columns, 8, 8, 20);

                        // Worksheet: Columns + SheetData + MergeCells (порядок важен)
                        MergeCells mergeCells = new MergeCells();
                        worksheetPart.Worksheet = new Worksheet(columns, sheetData, mergeCells);

                        Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
                        Sheet sheet = new Sheet()
                        {
                            Id = workbookPart.GetIdOfPart(worksheetPart),
                            SheetId = 1,
                            Name = "RoomBook"
                        };
                        sheets.Append(sheet);

                        var sharedCache = new Dictionary<string, int>(StringComparer.Ordinal);

                        // ----- Header rows -----
                        // Row 1 merged A1:H1
                        SetText(sheetData, sstPart, sharedCache, 1, 1, "Таблица вид 2", STYLE_CENTER10);
                        Merge(mergeCells, 1, 1, 1, 8);

                        // Row 2 merged A2:H2
                        SetText(sheetData, sstPart, sharedCache, 2, 1, "Румбук - Спецификация помещений", STYLE_CENTER14_BOLD);
                        Merge(mergeCells, 2, 1, 2, 8);

                        // Row 3 merged A3:H3 (left)
                        SetText(sheetData, sstPart, sharedCache, 3, 1, "Ссылки на листы документации", STYLE_LEFT10);
                        Merge(mergeCells, 3, 1, 3, 8);

                        // Row 4 column headers (with border)
                        string[] headers = new string[]
                        {
                            "Номер помещения", "Имя помещения", "Тип элемента", "Марка элемента",
                            "Наименование элемента", "Ед. изм", "Кол-во", "Примечание"
                        };
                        for (int i = 0; i < headers.Length; i++)
                            SetText(sheetData, sstPart, sharedCache, 4, i + 1, headers[i], STYLE_CENTER10_BOLD_BORDER);

                        int row = 5;

                        foreach (Room room in roomList)
                        {
                            step++;
                            roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Value = step);
                            roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.label_ItemName.Content = room.Name);

                            int startRow = row;

                            // Полы
                            List<Floor> floorList = new FilteredElementCollector(doc)
                                .OfCategory(BuiltInCategory.OST_Floors)
                                .OfClass(typeof(Floor))
                                .WhereElementIsNotElementType()
                                .Cast<Floor>()
                                .Where(f => f.FloorType.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL) != null)
                                .Where(f =>
                                {
                                    var m = GetStringParam(f.FloorType, BuiltInParameter.ALL_MODEL_MODEL);
                                    return m == "Пол" || m == "Полы";
                                })
                                .Where(f => f.get_Parameter(roombookRoomNumber) != null)
                                .Where(f => GetStringParam(f, roombookRoomNumber) == room.Number)
                                .OrderBy(f => GetStringParam(f.FloorType, BuiltInParameter.WINDOW_TYPE_ID), new AlphanumComparatorFastString())
                                .ToList();

                            // Стены
                            List<Wall> wallList = new FilteredElementCollector(doc)
                                .OfCategory(BuiltInCategory.OST_Walls)
                                .OfClass(typeof(Wall))
                                .WhereElementIsNotElementType()
                                .Cast<Wall>()
                                .Where(w => w.WallType.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL) != null &&
                                    !string.IsNullOrWhiteSpace(GetStringParam(w.WallType, BuiltInParameter.ALL_MODEL_MODEL)))
                                .Where(w => GetStringParam(w.WallType, BuiltInParameter.ALL_MODEL_MODEL) == "Отделка стен")
                                .Where(w => w.get_Parameter(roombookRoomNumber) != null)
                                .Where(w => GetStringParam(w, roombookRoomNumber) == room.Number)
                                .OrderBy(w => GetStringParam(w.WallType, BuiltInParameter.WINDOW_TYPE_ID), new AlphanumComparatorFastString())
                                .ToList();

                            // Потолки
                            List<Ceiling> ceilingList = new FilteredElementCollector(doc)
                                .OfCategory(BuiltInCategory.OST_Ceilings)
                                .OfClass(typeof(Ceiling))
                                .WhereElementIsNotElementType()
                                .Cast<Ceiling>()
                                .Where(c =>
                                {
                                    var model = GetStringParam(doc.GetElement(c.GetTypeId()), BuiltInParameter.ALL_MODEL_MODEL);
                                    return model == "Потолок" || model == "Потолки";
                                })
                                .Where(c => c.get_Parameter(roombookRoomNumber) != null)
                                .Where(c => GetStringParam(c, roombookRoomNumber) == room.Number)
                                .OrderBy(c => GetStringParam(doc.GetElement(c.GetTypeId()), BuiltInParameter.WINDOW_TYPE_ID), new AlphanumComparatorFastString())
                                .ToList();

                            if (floorList.Count == 0 && wallList.Count == 0 && ceilingList.Count == 0)
                                continue;

                            // Полы (уникальные типы)
                            List<FloorType> floorTypesList = new List<FloorType>();
                            List<ElementId> floorTypesIdList = new List<ElementId>();
                            foreach (Floor floor in floorList)
                            {
                                if (!floorTypesIdList.Contains(floor.FloorType.Id))
                                {
                                    floorTypesList.Add(floor.FloorType);
                                    floorTypesIdList.Add(floor.FloorType.Id);
                                }
                            }
                            floorTypesList = floorTypesList
                                .OrderBy(ft => GetStringParam(ft, BuiltInParameter.WINDOW_TYPE_ID), new AlphanumComparatorFastString())
                                .ToList();

                            foreach (FloorType floorType in floorTypesList)
                            {
                                double floorArea = 0;
                                List<Floor> tmpFloorList = floorList.Where(w => w.FloorType.Id == floorType.Id).ToList();
                                foreach (Floor floor in tmpFloorList)
                                {
#if R2019 || R2020 || R2021
                                    floorArea += UnitUtils.ConvertFromInternalUnits(floor.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsDouble(), DisplayUnitType.DUT_SQUARE_METERS);
#else
                                    floorArea += UnitUtils.ConvertFromInternalUnits(floor.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsDouble(), UnitTypeId.SquareMeters);
#endif
                                }

                                SetText(sheetData, sstPart, sharedCache, row, 1, room.Number, STYLE_CENTER10_BORDER);
                                SetText(sheetData, sstPart, sharedCache, row, 2, GetRoomName(room), STYLE_CENTER10_BORDER);
                                SetText(sheetData, sstPart, sharedCache, row, 3, "Пол", STYLE_CENTER10_BORDER);
                                SetText(sheetData, sstPart, sharedCache, row, 4, GetStringParam(floorType, BuiltInParameter.WINDOW_TYPE_ID), STYLE_CENTER10_BORDER);
                                SetText(sheetData, sstPart, sharedCache, row, 5, GetElementDescription(floorType, elemData), STYLE_LEFT10_BORDER);
                                SetText(sheetData, sstPart, sharedCache, row, 6, "м2", STYLE_CENTER10_BORDER);
                                SetNumber(sheetData, row, 7, Math.Round(floorArea, 2), STYLE_CENTER10_BORDER);
                                SetText(sheetData, sstPart, sharedCache, row, 8, "", STYLE_CENTER10_BORDER);
                                row++;
                            }

                            // Стены (уникальные типы)
                            List<WallType> wallTypesList = new List<WallType>();
                            List<ElementId> wallTypesIdList = new List<ElementId>();
                            foreach (Wall wall in wallList)
                            {
                                if (!wallTypesIdList.Contains(wall.WallType.Id))
                                {
                                    wallTypesList.Add(wall.WallType);
                                    wallTypesIdList.Add(wall.WallType.Id);
                                }
                            }
                            wallTypesList = wallTypesList
                                .OrderBy(wt => GetStringParam(wt, BuiltInParameter.WINDOW_TYPE_ID), new AlphanumComparatorFastString())
                                .ToList();

                            foreach (WallType wallType in wallTypesList)
                            {
                                double wallArea = 0;
                                List<Wall> tmpWallList = wallList.Where(w => w.WallType.Id == wallType.Id).ToList();
                                foreach (Wall wall in tmpWallList)
                                {
#if R2019 || R2020 || R2021
                                    wallArea += UnitUtils.ConvertFromInternalUnits(wall.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsDouble(), DisplayUnitType.DUT_SQUARE_METERS);
#else
                                    wallArea += UnitUtils.ConvertFromInternalUnits(wall.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsDouble(), UnitTypeId.SquareMeters);
#endif
                                }

                                SetText(sheetData, sstPart, sharedCache, row, 1, room.Number, STYLE_CENTER10_BORDER);
                                SetText(sheetData, sstPart, sharedCache, row, 2, GetRoomName(room), STYLE_CENTER10_BORDER);
                                SetText(sheetData, sstPart, sharedCache, row, 3, "Отделка\r\nстен", STYLE_CENTER10_BORDER);
                                SetText(sheetData, sstPart, sharedCache, row, 4, GetStringParam(wallType, BuiltInParameter.WINDOW_TYPE_ID), STYLE_CENTER10_BORDER);
                                SetText(sheetData, sstPart, sharedCache, row, 5, GetElementDescription(wallType, elemData), STYLE_LEFT10_BORDER);
                                SetText(sheetData, sstPart, sharedCache, row, 6, "м2", STYLE_CENTER10_BORDER);
                                SetNumber(sheetData, row, 7, Math.Round(wallArea, 2), STYLE_CENTER10_BORDER);
                                SetText(sheetData, sstPart, sharedCache, row, 8, "", STYLE_CENTER10_BORDER);
                                row++;
                            }

                            // Потолки (уникальные типы)
                            List<CeilingType> ceilingTypesList = new List<CeilingType>();
                            List<ElementId> ceilingTypesIdList = new List<ElementId>();
                            foreach (Ceiling ceiling in ceilingList)
                            {
                                if (!ceilingTypesIdList.Contains(ceiling.GetTypeId()))
                                {
                                    ceilingTypesList.Add(doc.GetElement(ceiling.GetTypeId()) as CeilingType);
                                    ceilingTypesIdList.Add(ceiling.GetTypeId());
                                }
                            }
                            ceilingTypesList = ceilingTypesList
                                .Where(ct => ct != null)
                                .OrderBy(ct => GetStringParam(ct, BuiltInParameter.WINDOW_TYPE_ID), new AlphanumComparatorFastString())
                                .ToList();

                            foreach (CeilingType ceilingType in ceilingTypesList)
                            {
                                double ceilingArea = 0;
                                List<Ceiling> tmpCeilingList = ceilingList.Where(c => c.GetTypeId() == ceilingType.Id).ToList();
                                foreach (Ceiling ceiling in tmpCeilingList)
                                {
#if R2019 || R2020 || R2021
                                    ceilingArea += UnitUtils.ConvertFromInternalUnits(ceiling.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsDouble(), DisplayUnitType.DUT_SQUARE_METERS);
#else
                                    ceilingArea += UnitUtils.ConvertFromInternalUnits(ceiling.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsDouble(), UnitTypeId.SquareMeters);
#endif
                                }

                                SetText(sheetData, sstPart, sharedCache, row, 1, room.Number, STYLE_CENTER10_BORDER);
                                SetText(sheetData, sstPart, sharedCache, row, 2, GetRoomName(room), STYLE_CENTER10_BORDER);
                                SetText(sheetData, sstPart, sharedCache, row, 3, "Потолок", STYLE_CENTER10_BORDER);
                                SetText(sheetData, sstPart, sharedCache, row, 4, GetStringParam(ceilingType, BuiltInParameter.WINDOW_TYPE_ID), STYLE_CENTER10_BORDER);
                                SetText(sheetData, sstPart, sharedCache, row, 5, GetElementDescription(ceilingType, elemData), STYLE_LEFT10_BORDER);
                                SetText(sheetData, sstPart, sharedCache, row, 6, "м2", STYLE_CENTER10_BORDER);
                                SetNumber(sheetData, row, 7, Math.Round(ceilingArea, 2), STYLE_CENTER10_BORDER);
                                SetText(sheetData, sstPart, sharedCache, row, 8, "", STYLE_CENTER10_BORDER);
                                row++;
                            }

                            int endRow = row - 1;
                            if (endRow >= startRow)
                            {
                                Merge(mergeCells, startRow, 1, endRow, 1);
                                Merge(mergeCells, startRow, 2, endRow, 2);
                            }
                        }

                        sstPart.SharedStringTable.Save();
                        worksheetPart.Worksheet.Save();
                        workbookPart.Workbook.Save();
                    }

                    roomBookToExcelProgressBarWPF.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.Close());
                }
                catch (Exception ex)
                {
                    roomBookToExcelProgressBarWPF.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.Close());
                    string errorMessage = "Error: " + ex.Message + " Line: " + ex.Source;
                    TaskDialog.Show("Revit", errorMessage);
                    return Result.Cancelled;
                }
            }

            else if (exportOptionName == "rbt_FloorFinishByCombinationInRoom")
            {
                StartProgressWindow(roomList.Count);
                int step = 0;

                try
                {
                    List<ItemFloorFinishByRoom> itemFloorFinishByRoomList = new List<ItemFloorFinishByRoom>();
                    foreach (Room room in roomList)
                    {
                        step++;
                        roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Value = step);
                        roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() =>
                            roomBookToExcelProgressBarWPF.label_ItemName.Content = $"Сбор данных об отделке пола. Шаг {step} из {roomList.Count}");

                        ItemFloorFinishByRoom itemFloorFinishByRoom = new ItemFloorFinishByRoom();
                        itemFloorFinishByRoom.RoomNumber = room.Number;
                        itemFloorFinishByRoom.RoomName = GetRoomName(room);

                        List<Floor> floorList = new FilteredElementCollector(doc)
                            .OfCategory(BuiltInCategory.OST_Floors)
                            .OfClass(typeof(Floor))
                            .WhereElementIsNotElementType()
                            .Cast<Floor>()
                            .Where(w => w.FloorType.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL) != null)
                            .Where(f => {
                                var model = GetStringParam(f.FloorType, BuiltInParameter.ALL_MODEL_MODEL);
                                return model == "Пол" || model == "Полы";
                            })
                            .Where(w => w.get_Parameter(roombookRoomNumber) != null)
                            .Where(w => GetStringParam(w, roombookRoomNumber) == room.Number)
                            .OrderBy(w => GetStringParam(w.FloorType, BuiltInParameter.WINDOW_TYPE_ID), new AlphanumComparatorFastString())
                            .ToList();

                        List<FloorType> floorTypesList = new List<FloorType>();
                        List<ElementId> floorTypesIdList = new List<ElementId>();
                        foreach (Floor floor in floorList)
                        {
                            if (!floorTypesIdList.Contains(floor.FloorType.Id))
                            {
                                floorTypesList.Add(floor.FloorType);
                                floorTypesIdList.Add(floor.FloorType.Id);
                            }
                        }
                        itemFloorFinishByRoom.FloorTypesList = floorTypesList
                            .OrderBy(wt => GetStringParam(wt, BuiltInParameter.WINDOW_TYPE_ID), new AlphanumComparatorFastString())
                            .ToList();
                        itemFloorFinishByRoomList.Add(itemFloorFinishByRoom);
                    }

                    List<ItemFloorFinishByRoom> uniqueFloorFinishSet = itemFloorFinishByRoomList
                        .Distinct(new ItemFloorFinishByRoomComparer())
                        .ToList();

                    step = 0;
                    roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Minimum = 0);
                    roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Maximum = uniqueFloorFinishSet.Count);

                    List<ItemFloorFinishByRoomExcelString> itemFloorFinishByRoomExcelStringList = new List<ItemFloorFinishByRoomExcelString>();
                    foreach (ItemFloorFinishByRoom uniqueFloorFinish in uniqueFloorFinishSet)
                    {
                        step++;
                        roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Value = step);
                        roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() =>
                            roomBookToExcelProgressBarWPF.label_ItemName.Content = $"Обработка сочетаний отделок. Шаг {step} из {uniqueFloorFinishSet.Count}");

                        ItemFloorFinishByRoomExcelString itemFloorFinishByRoomExcelString = new ItemFloorFinishByRoomExcelString();
                        itemFloorFinishByRoomExcelString.ItemData = new Dictionary<string, double>();
                        List<ItemFloorFinishByRoom> tmpItemFloorFinishList = itemFloorFinishByRoomList
                            .Where(i => i.Equals(uniqueFloorFinish))
                            .OrderBy(i => i.RoomNumber, new AlphanumComparatorFastString())
                            .ToList();

                        List<string> roomNumbersList = new List<string>();
                        List<string> roomNamesList = new List<string>();

                        foreach (ItemFloorFinishByRoom tmpItemFloorFinish in tmpItemFloorFinishList)
                        {
                            if (!roomNumbersList.Contains(tmpItemFloorFinish.RoomNumber))
                                roomNumbersList.Add(tmpItemFloorFinish.RoomNumber);

                            if (!roomNamesList.Contains(tmpItemFloorFinish.RoomName))
                                roomNamesList.Add(tmpItemFloorFinish.RoomName);

                            foreach (FloorType floorType in tmpItemFloorFinish.FloorTypesList)
                            {
                                List<Floor> tmpFloorList = new FilteredElementCollector(doc)
                                    .OfCategory(BuiltInCategory.OST_Floors)
                                    .OfClass(typeof(Floor))
                                    .WhereElementIsNotElementType()
                                    .Cast<Floor>()
                                    .Where(w => w.FloorType.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL) != null)
                                    .Where(f => {
                                        var model = GetStringParam(f.FloorType, BuiltInParameter.ALL_MODEL_MODEL);
                                        return model == "Пол" || model == "Полы";
                                    })
                                    .Where(w => w.get_Parameter(roombookRoomNumber) != null)
                                    .Where(w => GetStringParam(w, roombookRoomNumber) == tmpItemFloorFinish.RoomNumber)
                                    .Where(w => w.FloorType.Id == floorType.Id)
                                    .ToList();

                                double floorArea = 0;
                                foreach (Floor floor in tmpFloorList)
                                {
#if R2019 || R2020 || R2021
                                    floorArea += UnitUtils.ConvertFromInternalUnits(floor.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsDouble(), DisplayUnitType.DUT_SQUARE_METERS);
#else
                                    floorArea += UnitUtils.ConvertFromInternalUnits(floor.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsDouble(), UnitTypeId.SquareMeters);
#endif
                                }

                                string floorTypeDisc = GetStringParam(floorType, elemData);
                                if (string.IsNullOrWhiteSpace(floorTypeDisc)) floorTypeDisc = "(без имени типа)";
                                if (itemFloorFinishByRoomExcelString.ItemData.ContainsKey(floorTypeDisc))
                                    itemFloorFinishByRoomExcelString.ItemData[floorTypeDisc] += Math.Round(floorArea, 2);
                                else
                                    itemFloorFinishByRoomExcelString.ItemData.Add(floorTypeDisc, Math.Round(floorArea, 2));
                            }
                        }

                        itemFloorFinishByRoomExcelString.RoomNumber = string.Join(", ", roomNumbersList.OrderBy(n => n, new AlphanumComparatorFastString()));
                        itemFloorFinishByRoomExcelString.RoomName = string.Join(", ", roomNamesList.OrderBy(n => n, new AlphanumComparatorFastString()));
                        itemFloorFinishByRoomExcelStringList.Add(itemFloorFinishByRoomExcelString);
                    }
                    roomBookToExcelProgressBarWPF.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.Close());

                    var saveDialog = new System.Windows.Forms.SaveFileDialog();
                    saveDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
                    var dlgRes = saveDialog.ShowDialog();
                    if (dlgRes != System.Windows.Forms.DialogResult.OK)
                    {
                        return Result.Cancelled;
                    }

                    string filePath = saveDialog.FileName;
                    int itemDataCnt = itemFloorFinishByRoomExcelStringList.Any() ? itemFloorFinishByRoomExcelStringList.Max(i => i.ItemData.Count) * 2 : 2;
                    uint lastColumn = (uint)(itemDataCnt + 3);

                    using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
                    {
                        WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();
                        workbookPart.Workbook = new Workbook();

                        WorkbookStylesPart stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
                        stylesPart.Stylesheet = BuildStylesheet_ISOCPEUR();
                        stylesPart.Stylesheet.Save();

                        SharedStringTablePart sstPart = workbookPart.AddNewPart<SharedStringTablePart>();
                        sstPart.SharedStringTable = new SharedStringTable();

                        WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                        SheetData sheetData = new SheetData();
                        Columns columns = new Columns();
                        AddColumn(columns, 1, 1, 60);
                        AddColumn(columns, 2, 2, 120);
                        for (uint i = 3; i <= itemDataCnt + 2; i += 2)
                        {
                            AddColumn(columns, i, i, 65);
                            AddColumn(columns, i + 1, i + 1, 15);
                        }
                        AddColumn(columns, lastColumn, lastColumn, 20);

                        MergeCells mergeCells = new MergeCells();
                        worksheetPart.Worksheet = new Worksheet(columns, sheetData, mergeCells);

                        Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
                        Sheet sheet = new Sheet()
                        {
                            Id = workbookPart.GetIdOfPart(worksheetPart),
                            SheetId = 1,
                            Name = "FloorFinish"
                        };
                        sheets.Append(sheet);

                        var sharedCache = new Dictionary<string, int>(StringComparer.Ordinal);

                        SetText(sheetData, sstPart, sharedCache, 1, 1, "Ведомость отделки пола", STYLE_CENTER10_BOLD);
                        Merge(mergeCells, 1, 1, 1, (int)lastColumn);

                        SetText(sheetData, sstPart, sharedCache, 2, 1, "Номера помещений", STYLE_CENTER10_BORDER);
                        Merge(mergeCells, 2, 1, 3, 1);

                        SetText(sheetData, sstPart, sharedCache, 2, 2, "Наименования помещений", STYLE_CENTER10_BORDER);
                        Merge(mergeCells, 2, 2, 3, 2);

                        SetText(sheetData, sstPart, sharedCache, 2, 3, "Типы отделки помещений", STYLE_CENTER10_BORDER);
                        Merge(mergeCells, 2, 3, 2, itemDataCnt + 2);

                        SetText(sheetData, sstPart, sharedCache, 2, itemDataCnt + 3, "Примечание", STYLE_CENTER10_BORDER);
                        Merge(mergeCells, 2, itemDataCnt + 3, 3, itemDataCnt + 3);

                        int typeCnt = 1;
                        for (int i = 3; i <= itemDataCnt + 2; i += 2)
                        {
                            SetText(sheetData, sstPart, sharedCache, 3, i, $"Отделка пола тип {typeCnt}", STYLE_CENTER10_BORDER);
                            SetText(sheetData, sstPart, sharedCache, 3, i + 1, "Площ. м2", STYLE_CENTER10_BORDER);
                            typeCnt++;
                        }

                        int row = 4;
                        foreach (ItemFloorFinishByRoomExcelString item in itemFloorFinishByRoomExcelStringList)
                        {
                            SetText(sheetData, sstPart, sharedCache, row, 1, item.RoomNumber, STYLE_CENTER10_BORDER);
                            SetText(sheetData, sstPart, sharedCache, row, 2, item.RoomName, STYLE_CENTER10_BORDER);

                            int j = 0;
                            foreach (KeyValuePair<string, double> pair in item.ItemData)
                            {
                                SetText(sheetData, sstPart, sharedCache, row, j * 2 + 3, pair.Key, STYLE_LEFT10_BORDER);
                                SetNumber(sheetData, row, j * 2 + 4, pair.Value, STYLE_CENTER10_BORDER);
                                j++;
                            }

                            SetText(sheetData, sstPart, sharedCache, row, itemDataCnt + 3, "", STYLE_CENTER10_BORDER);
                            row++;
                        }

                        sstPart.SharedStringTable.Save();
                        worksheetPart.Worksheet.Save();
                        workbookPart.Workbook.Save();
                    }
                }
                catch (Exception ex)
                {
                    roomBookToExcelProgressBarWPF.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.Close());
                    string errorMessage = "Error: " + ex.Message + " Line: " + ex.Source;
                    TaskDialog.Show("Revit", errorMessage);
                    return Result.Cancelled;
                }
            }

            else if (exportOptionName == "rbt_WallFinishByCombinationInRoom")
            {
                StartProgressWindow(roomList.Count);
                int step = 0;

                try
                {
                    var combos = new List<ItemWallFinishByRoom>();
                    foreach (Room room in roomList)
                    {
                        step++;
                        roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Value = step);
                        roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.label_ItemName.Content = $"Сбор данных. Шаг {step} из {roomList.Count}");
                        var item = new ItemWallFinishByRoom
                        {
                            RoomNumber = room.Number,
                            RoomName = GetRoomName(room)
                        };

                        var finishes = new FilteredElementCollector(doc)
                            .OfCategory(BuiltInCategory.OST_Walls)
                            .OfClass(typeof(Wall))
                            .WhereElementIsNotElementType()
                            .Cast<Wall>()
                            .Where(w =>
                                GetStringParam(w.WallType, BuiltInParameter.ALL_MODEL_MODEL).StartsWith("Отделка стен") &&
                                w.get_Parameter(roombookRoomNumber) != null &&
                                GetStringParam(w, roombookRoomNumber) == room.Number)
                            .ToList();

                        item.WallFinishes.AddRange(finishes
                            .Where(w => !GetStringParam(w.WallType, BuiltInParameter.ALL_MODEL_MODEL).Contains("Колонны"))
                            .Select(w => w.WallType)
                            .Distinct(new WallTypeIdComparer())
                            .OrderBy(t => t.Name, StringComparer.Ordinal));

                        item.ColumnFinishes.AddRange(finishes
                            .Where(w => GetStringParam(w.WallType, BuiltInParameter.ALL_MODEL_MODEL).Contains("Колонны"))
                            .Select(w => w.WallType)
                            .Distinct(new WallTypeIdComparer())
                            .OrderBy(t => t.Name, StringComparer.Ordinal));

                        combos.Add(item);
                    }

                    var uniqueCombos = combos.Distinct(new ItemWallFinishByRoomComparer()).OrderBy(c => c.RoomNumber, new AlphanumComparatorFastString()).ToList();
                    step = 0;
                    roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Maximum = uniqueCombos.Count);
                    var excelRows = new List<ItemWallFinishByRoomExcelString>();

                    foreach (var combo in uniqueCombos)
                    {
                        step++;
                        roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Value = step);
                        roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.label_ItemName.Content = $"Комбинация {step} из {uniqueCombos.Count}");

                        var rowData = new ItemWallFinishByRoomExcelString();
                        var rooms = combos.Where(c => c.Equals(combo)).ToList();
                        rowData.RoomNumber = string.Join(", ", rooms.Select(r => r.RoomNumber).Distinct().OrderBy(s => s, new AlphanumComparatorFastString()));
                        rowData.RoomName = string.Join(", ", rooms.Select(r => r.RoomName).Distinct(StringComparer.Ordinal).OrderBy(s => s, new AlphanumComparatorFastString()));

                        foreach (var room in rooms)
                        {
                            AccumulateWallAreas(doc, room.WallFinishes, room.RoomNumber, rowData.ItemData, elemData, roombookRoomNumber, false, false);
                            AccumulateWallAreas(doc, room.ColumnFinishes, room.RoomNumber, rowData.ItemData, elemData, roombookRoomNumber, true, true);
                        }

                        if (rowData.ItemData.Count > 0)
                            excelRows.Add(rowData);
                    }
                    roomBookToExcelProgressBarWPF.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.Close());

                    var saveDialog = new System.Windows.Forms.SaveFileDialog();
                    saveDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
                    if (saveDialog.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                        return Result.Cancelled;

                    using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(saveDialog.FileName, SpreadsheetDocumentType.Workbook))
                    {
                        WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();
                        workbookPart.Workbook = new Workbook();
                        WorkbookStylesPart stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
                        stylesPart.Stylesheet = BuildStylesheet_ISOCPEUR();
                        stylesPart.Stylesheet.Save();
                        SharedStringTablePart sstPart = workbookPart.AddNewPart<SharedStringTablePart>();
                        sstPart.SharedStringTable = new SharedStringTable();
                        WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                        SheetData sheetData = new SheetData();
                        Columns columns = new Columns();
                        AddColumn(columns, 1, 1, 60);
                        AddColumn(columns, 2, 2, 60);
                        AddColumn(columns, 3, 3, 15);
                        AddColumn(columns, 4, 4, 60);
                        AddColumn(columns, 5, 5, 15);
                        AddColumn(columns, 6, 6, 60);
                        AddColumn(columns, 7, 7, 15);
                        AddColumn(columns, 8, 8, 17);
                        MergeCells mergeCells = new MergeCells();
                        worksheetPart.Worksheet = new Worksheet(columns, sheetData, mergeCells);
                        Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
                        sheets.Append(new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Отделка стен" });
                        var sharedCache = new Dictionary<string, int>(StringComparer.Ordinal);

                        SetText(sheetData, sstPart, sharedCache, 1, 1, "Ведомость отделки", STYLE_CENTER10_BOLD);
                        Merge(mergeCells, 1, 1, 1, 8);
                        SetText(sheetData, sstPart, sharedCache, 2, 1, "Номера помещений", STYLE_CENTER10_BOLD_BORDER);
                        Merge(mergeCells, 2, 1, 3, 1);
                        SetText(sheetData, sstPart, sharedCache, 2, 2, "Наименования помещений", STYLE_CENTER10_BOLD_BORDER);
                        Merge(mergeCells, 2, 2, 3, 2);
                        SetText(sheetData, sstPart, sharedCache, 2, 3, "Типы отделки помещений", STYLE_CENTER10_BOLD_BORDER);
                        Merge(mergeCells, 2, 3, 2, 7);
                        SetText(sheetData, sstPart, sharedCache, 2, 8, "Примечания", STYLE_CENTER10_BOLD_BORDER);
                        Merge(mergeCells, 2, 8, 3, 8);
                        SetText(sheetData, sstPart, sharedCache, 3, 3, "Отделка потолка", STYLE_CENTER10_BOLD_BORDER);
                        SetText(sheetData, sstPart, sharedCache, 3, 4, "Отделка стен или перегородок", STYLE_CENTER10_BOLD_BORDER);
                        SetText(sheetData, sstPart, sharedCache, 3, 5, "Площ. м2", STYLE_CENTER10_BOLD_BORDER);
                        SetText(sheetData, sstPart, sharedCache, 3, 6, "Отделка колонн", STYLE_CENTER10_BOLD_BORDER);
                        SetText(sheetData, sstPart, sharedCache, 3, 7, "Площ. м2", STYLE_CENTER10_BOLD_BORDER);

                        const string ceilingNote = "см. ведомость отделки потолков на листе";
                        int row = 4;
                        foreach (var item in excelRows)
                        {
                            var walls = item.ItemData.Where(k => !k.Key.StartsWith("Колонны — ")).OrderBy(k => k.Key).ToList();
                            var cols = item.ItemData.Where(k => k.Key.StartsWith("Колонны — ")).OrderBy(k => k.Key).Select(kv => new KeyValuePair<string, double>(kv.Key.Replace("Колонны — ", ""), kv.Value)).ToList();
                            int rows = Math.Max(walls.Count, cols.Count);
                            if (rows == 0) continue;
                            int startRow = row;
                            int endRow = row + rows - 1;

                            for (int i = 0; i < rows; i++)
                            {
                                int curRow = row + i;
                                SetText(sheetData, sstPart, sharedCache, curRow, 1, item.RoomNumber, STYLE_CENTER10_BORDER);
                                SetText(sheetData, sstPart, sharedCache, curRow, 2, item.RoomName, STYLE_CENTER10_BORDER);
                                SetText(sheetData, sstPart, sharedCache, curRow, 3, ceilingNote, STYLE_CENTER10_BORDER);
                                SetText(sheetData, sstPart, sharedCache, curRow, 4, i < walls.Count ? walls[i].Key : "", STYLE_LEFT10_BORDER);
                                if (i < walls.Count) SetNumber(sheetData, curRow, 5, walls[i].Value, STYLE_CENTER10_BORDER); else SetText(sheetData, sstPart, sharedCache, curRow, 5, "", STYLE_CENTER10_BORDER);
                                SetText(sheetData, sstPart, sharedCache, curRow, 6, i < cols.Count ? cols[i].Key : "", STYLE_LEFT10_BORDER);
                                if (i < cols.Count) SetNumber(sheetData, curRow, 7, cols[i].Value, STYLE_CENTER10_BORDER); else SetText(sheetData, sstPart, sharedCache, curRow, 7, "", STYLE_CENTER10_BORDER);
                                SetText(sheetData, sstPart, sharedCache, curRow, 8, "", STYLE_CENTER10_BORDER);
                            }

                            Merge(mergeCells, startRow, 1, endRow, 1);
                            Merge(mergeCells, startRow, 2, endRow, 2);
                            Merge(mergeCells, startRow, 3, endRow, 3);
                            Merge(mergeCells, startRow, 8, endRow, 8);
                            row = endRow + 1;
                        }

                        sstPart.SharedStringTable.Save();
                        worksheetPart.Worksheet.Save();
                        workbookPart.Workbook.Save();
                    }
                }
                catch (Exception ex)
                {
                    roomBookToExcelProgressBarWPF.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.Close());
                    TaskDialog.Show("Revit", "Error: " + ex.Message);
                    return Result.Cancelled;
                }
            }
            else if (exportOptionName == "rbt_CeilingFinishByCombinationInRoom")
            {
                StartProgressWindow(roomList.Count);
                int step = 0;

                try
                {
                    List<ItemCeilingFinishByRoom> itemCeilingFinishByRoomList = new List<ItemCeilingFinishByRoom>();
                    foreach (Room room in roomList)
                    {
                        step++;
                        roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Value = step);
                        roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.label_ItemName.Content = $"Сбор данных об отделке потолка. Шаг {step} из {roomList.Count}");

                        ItemCeilingFinishByRoom itemCeilingFinishByRoom = new ItemCeilingFinishByRoom();
                        itemCeilingFinishByRoom.RoomNumber = room.Number;
                        itemCeilingFinishByRoom.RoomName = GetRoomName(room);

                        List<Ceiling> ceilingList = new FilteredElementCollector(doc)
                            .OfCategory(BuiltInCategory.OST_Ceilings)
                            .OfClass(typeof(Ceiling))
                            .WhereElementIsNotElementType()
                            .Cast<Ceiling>()
                            .Where(c =>
                            {
                                var model = GetStringParam(doc.GetElement(c.GetTypeId()), BuiltInParameter.ALL_MODEL_MODEL);
                                return model == "Потолок" || model == "Потолки";
                            })
                            .Where(c => c.get_Parameter(roombookRoomNumber) != null)
                            .Where(c => GetStringParam(c, roombookRoomNumber) == room.Number)
                            .OrderBy(c => GetStringParam(doc.GetElement(c.GetTypeId()), BuiltInParameter.WINDOW_TYPE_ID), new AlphanumComparatorFastString())
                            .ToList();

                        List<CeilingType> ceilingTypesList = new List<CeilingType>();
                        List<ElementId> ceilingTypesIdList = new List<ElementId>();
                        foreach (Ceiling ceiling in ceilingList)
                        {
                            if (!ceilingTypesIdList.Contains(ceiling.GetTypeId()))
                            {
                                ceilingTypesList.Add(doc.GetElement(ceiling.GetTypeId()) as CeilingType);
                                ceilingTypesIdList.Add(ceiling.GetTypeId());
                            }
                        }
                        itemCeilingFinishByRoom.CeilingTypesList = ceilingTypesList.Where(ct => ct != null)
                            .OrderBy(wt => GetStringParam(wt, BuiltInParameter.WINDOW_TYPE_ID), new AlphanumComparatorFastString())
                            .ToList();
                        itemCeilingFinishByRoomList.Add(itemCeilingFinishByRoom);
                    }

                    var uniqueCeilingFinishSet = itemCeilingFinishByRoomList.Distinct(new ItemCeilingFinishByRoomComparer()).ToList();
                    step = 0;
                    roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Maximum = uniqueCeilingFinishSet.Count);
                    var excelRows = new List<ItemCeilingFinishByRoomExcelString>();

                    foreach (var uniqueFinish in uniqueCeilingFinishSet)
                    {
                        step++;
                        roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Value = step);
                        roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.label_ItemName.Content = $"Обработка сочетаний отделок. Шаг {step} из {uniqueCeilingFinishSet.Count}");

                        var rowData = new ItemCeilingFinishByRoomExcelString();
                        rowData.ItemData = new Dictionary<string, double>();
                        var rooms = itemCeilingFinishByRoomList.Where(i => i.Equals(uniqueFinish)).OrderBy(i => i.RoomNumber, new AlphanumComparatorFastString()).ToList();
                        rowData.RoomNumber = string.Join(", ", rooms.Select(r => r.RoomNumber).Distinct().OrderBy(n => n, new AlphanumComparatorFastString()));
                        rowData.RoomName = string.Join(", ", rooms.Select(r => r.RoomName).Distinct().OrderBy(n => n, new AlphanumComparatorFastString()));
                        foreach (var room in rooms)
                            AccumulateCeilingAreas(doc, room.CeilingTypesList, room.RoomNumber, rowData.ItemData, elemData, roombookRoomNumber);
                        excelRows.Add(rowData);
                    }
                    roomBookToExcelProgressBarWPF.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.Close());

                    var saveDialog = new System.Windows.Forms.SaveFileDialog();
                    saveDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
                    if (saveDialog.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                        return Result.Cancelled;

                    int itemDataCnt = excelRows.Any() ? excelRows.Max(i => i.ItemData.Count) * 2 : 2;
                    uint lastColumn = (uint)(itemDataCnt + 3);
                    using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(saveDialog.FileName, SpreadsheetDocumentType.Workbook))
                    {
                        WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();
                        workbookPart.Workbook = new Workbook();
                        WorkbookStylesPart stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
                        stylesPart.Stylesheet = BuildStylesheet_ISOCPEUR();
                        stylesPart.Stylesheet.Save();
                        SharedStringTablePart sstPart = workbookPart.AddNewPart<SharedStringTablePart>();
                        sstPart.SharedStringTable = new SharedStringTable();
                        WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                        SheetData sheetData = new SheetData();
                        Columns columns = new Columns();
                        AddColumn(columns, 1, 1, 60); AddColumn(columns, 2, 2, 120);
                        for (uint i = 3; i <= itemDataCnt + 2; i += 2) { AddColumn(columns, i, i, 65); AddColumn(columns, i + 1, i + 1, 15); }
                        AddColumn(columns, lastColumn, lastColumn, 20);
                        MergeCells mergeCells = new MergeCells();
                        worksheetPart.Worksheet = new Worksheet(columns, sheetData, mergeCells);
                        Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
                        sheets.Append(new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "CeilingFinish" });
                        var sharedCache = new Dictionary<string, int>(StringComparer.Ordinal);

                        SetText(sheetData, sstPart, sharedCache, 1, 1, "Ведомость отделки потолка", STYLE_CENTER10_BOLD); Merge(mergeCells, 1, 1, 1, (int)lastColumn);
                        SetText(sheetData, sstPart, sharedCache, 2, 1, "Номера помещений", STYLE_CENTER10_BOLD_BORDER); Merge(mergeCells, 2, 1, 3, 1);
                        SetText(sheetData, sstPart, sharedCache, 2, 2, "Наименования помещений", STYLE_CENTER10_BOLD_BORDER); Merge(mergeCells, 2, 2, 3, 2);
                        SetText(sheetData, sstPart, sharedCache, 2, 3, "Типы отделки помещений", STYLE_CENTER10_BOLD_BORDER); Merge(mergeCells, 2, 3, 2, itemDataCnt + 2);
                        SetText(sheetData, sstPart, sharedCache, 2, itemDataCnt + 3, "Примечание", STYLE_CENTER10_BOLD_BORDER); Merge(mergeCells, 2, itemDataCnt + 3, 3, itemDataCnt + 3);
                        int typeCnt = 1;
                        for (int i = 3; i <= itemDataCnt + 2; i += 2)
                        {
                            SetText(sheetData, sstPart, sharedCache, 3, i, $"Отделка потолка тип {typeCnt}", STYLE_CENTER10_BOLD_BORDER);
                            SetText(sheetData, sstPart, sharedCache, 3, i + 1, "Площ. м2", STYLE_CENTER10_BOLD_BORDER);
                            typeCnt++;
                        }

                        int row = 4;
                        foreach (var item in excelRows)
                        {
                            SetText(sheetData, sstPart, sharedCache, row, 1, item.RoomNumber, STYLE_CENTER10_BORDER);
                            SetText(sheetData, sstPart, sharedCache, row, 2, item.RoomName, STYLE_CENTER10_BORDER);
                            int j = 0;
                            foreach (var pair in item.ItemData)
                            {
                                SetText(sheetData, sstPart, sharedCache, row, j * 2 + 3, pair.Key, STYLE_LEFT10_BORDER);
                                SetNumber(sheetData, row, j * 2 + 4, pair.Value, STYLE_CENTER10_BORDER);
                                j++;
                            }
                            SetText(sheetData, sstPart, sharedCache, row, itemDataCnt + 3, "", STYLE_CENTER10_BORDER);
                            row++;
                        }
                        sstPart.SharedStringTable.Save();
                        worksheetPart.Worksheet.Save();
                        workbookPart.Workbook.Save();
                    }
                }
                catch (Exception ex)
                {
                    roomBookToExcelProgressBarWPF.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.Close());
                    TaskDialog.Show("Revit", "Error: " + ex.Message);
                    return Result.Cancelled;
                }
            }
            else if (exportOptionName == "rbt_WallFinishByCombinationWithCeiling")
            {
                StartProgressWindow(roomList.Count);
                int step = 0;

                try
                {
                    var combos = new List<ItemWallFinishByRoomWithCeil>();
                    foreach (Room room in roomList)
                    {
                        step++;
                        roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Value = step);
                        roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.label_ItemName.Content = $"Сбор данных. Шаг {step} из {roomList.Count}");
                        var item = new ItemWallFinishByRoomWithCeil { RoomNumber = room.Number, RoomName = GetRoomName(room) };
                        var finishes = new FilteredElementCollector(doc)
                            .OfCategory(BuiltInCategory.OST_Walls)
                            .OfClass(typeof(Wall))
                            .WhereElementIsNotElementType()
                            .Cast<Wall>()
                            .Where(w =>
                                GetStringParam(w.WallType, BuiltInParameter.ALL_MODEL_MODEL).StartsWith("Отделка стен") &&
                                w.get_Parameter(roombookRoomNumber) != null &&
                                GetStringParam(w, roombookRoomNumber) == room.Number)
                            .ToList();
                        item.WallFinishes.AddRange(finishes.Where(w => !GetStringParam(w.WallType, BuiltInParameter.ALL_MODEL_MODEL).Contains("Колонны")).Select(w => w.WallType).Distinct(new WallTypeIdComparer()).OrderBy(t => t.Name, StringComparer.Ordinal));
                        item.ColumnFinishes.AddRange(finishes.Where(w => GetStringParam(w.WallType, BuiltInParameter.ALL_MODEL_MODEL).Contains("Колонны")).Select(w => w.WallType).Distinct(new WallTypeIdComparer()).OrderBy(t => t.Name, StringComparer.Ordinal));
                        item.CeilingFinishes.AddRange(new FilteredElementCollector(doc)
                            .OfCategory(BuiltInCategory.OST_Ceilings)
                            .OfClass(typeof(Ceiling))
                            .WhereElementIsNotElementType()
                            .Cast<Ceiling>()
                            .Where(c => c.get_Parameter(roombookRoomNumber) != null && GetStringParam(c, roombookRoomNumber) == room.Number)
                            .Select(c => doc.GetElement(c.GetTypeId()) as CeilingType)
                            .Where(ct => ct != null && ct.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL) != null && (GetStringParam(ct, BuiltInParameter.ALL_MODEL_MODEL) == "Потолок" || GetStringParam(ct, BuiltInParameter.ALL_MODEL_MODEL) == "Потолки"))
                            .Distinct(new ElementTypeIdComparer<CeilingType>())
                            .OrderBy(ct => ct.get_Parameter(BuiltInParameter.WINDOW_TYPE_ID)?.AsString() ?? ct.Name, new AlphanumComparatorFastString()));
                        combos.Add(item);
                    }
                    var uniqueCombos = combos.Distinct(new ItemWallFinishByRoomWithCeilComparer()).OrderBy(c => c.RoomNumber, new AlphanumComparatorFastString()).ToList();
                    roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Maximum = uniqueCombos.Count);
                    var excelRows = new List<ItemWallFinishByRoomWithCeilExcelRow>();

                    foreach (var combo in uniqueCombos)
                    {
                        step++;
                        roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Value = step);
                        roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.label_ItemName.Content = $"Комбинация {step} из {uniqueCombos.Count}");

                        var rowData = new ItemWallFinishByRoomWithCeilExcelRow();
                        var rooms = combos.Where(c => c.Equals(combo)).ToList();
                        rowData.RoomNumber = string.Join(", ", rooms.Select(r => r.RoomNumber).Distinct().OrderBy(s => s, new AlphanumComparatorFastString()));
                        rowData.RoomName = string.Join(", ", rooms.Select(r => r.RoomName).Distinct(StringComparer.Ordinal).OrderBy(s => s, new AlphanumComparatorFastString()));

                        foreach (var room in rooms)
                        {
                            AccumulateWallAreas(doc, room.WallFinishes, room.RoomNumber, rowData.WallData, elemData, roombookRoomNumber, false, false);
                            AccumulateWallAreas(doc, room.ColumnFinishes, room.RoomNumber, rowData.ColumnData, elemData, roombookRoomNumber, true, false);
                            AccumulateCeilingAreas(doc, combo.CeilingFinishes, room.RoomNumber, rowData.CeilingData, elemData, roombookRoomNumber);
                        }
                        if (rowData.WallData.Count > 0 || rowData.ColumnData.Count > 0 || rowData.CeilingData.Count > 0)
                            excelRows.Add(rowData);
                    }
                    roomBookToExcelProgressBarWPF.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.Close());

                    var saveDialog = new System.Windows.Forms.SaveFileDialog();
                    saveDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
                    if (saveDialog.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                        return Result.Cancelled;

                    using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(saveDialog.FileName, SpreadsheetDocumentType.Workbook))
                    {
                        WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();
                        workbookPart.Workbook = new Workbook();
                        WorkbookStylesPart stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
                        stylesPart.Stylesheet = BuildStylesheet_ISOCPEUR();
                        stylesPart.Stylesheet.Save();
                        SharedStringTablePart sstPart = workbookPart.AddNewPart<SharedStringTablePart>();
                        sstPart.SharedStringTable = new SharedStringTable();
                        WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                        SheetData sheetData = new SheetData();
                        Columns columns = new Columns();
                        AddColumn(columns, 1, 1, 50); AddColumn(columns, 2, 2, 50); AddColumn(columns, 3, 3, 45); AddColumn(columns, 4, 4, 15); AddColumn(columns, 5, 5, 45); AddColumn(columns, 6, 6, 15); AddColumn(columns, 7, 7, 50); AddColumn(columns, 8, 8, 15); AddColumn(columns, 9, 9, 16);
                        MergeCells mergeCells = new MergeCells();
                        worksheetPart.Worksheet = new Worksheet(columns, sheetData, mergeCells);
                        Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
                        sheets.Append(new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Стены и потолки" });
                        var sharedCache = new Dictionary<string, int>(StringComparer.Ordinal);

                        SetText(sheetData, sstPart, sharedCache, 1, 1, "Ведомость отделки", STYLE_CENTER10_BOLD); Merge(mergeCells, 1, 1, 1, 9);
                        SetText(sheetData, sstPart, sharedCache, 2, 1, "Номера помещений", STYLE_CENTER10_BOLD_BORDER); Merge(mergeCells, 2, 1, 3, 1);
                        SetText(sheetData, sstPart, sharedCache, 2, 2, "Наименования помещений", STYLE_CENTER10_BOLD_BORDER); Merge(mergeCells, 2, 2, 3, 2);
                        SetText(sheetData, sstPart, sharedCache, 2, 3, "Типы отделки помещений", STYLE_CENTER10_BOLD_BORDER); Merge(mergeCells, 2, 3, 2, 8);
                        SetText(sheetData, sstPart, sharedCache, 2, 9, "Примечания", STYLE_CENTER10_BOLD_BORDER); Merge(mergeCells, 2, 9, 3, 9);
                        SetText(sheetData, sstPart, sharedCache, 3, 3, "Отделка потолка", STYLE_CENTER10_BOLD_BORDER);
                        SetText(sheetData, sstPart, sharedCache, 3, 4, "Площ. м2", STYLE_CENTER10_BOLD_BORDER);
                        SetText(sheetData, sstPart, sharedCache, 3, 5, "Отделка стен или перегородок", STYLE_CENTER10_BOLD_BORDER);
                        SetText(sheetData, sstPart, sharedCache, 3, 6, "Площ. м2", STYLE_CENTER10_BOLD_BORDER);
                        SetText(sheetData, sstPart, sharedCache, 3, 7, "Отделка колонн", STYLE_CENTER10_BOLD_BORDER);
                        SetText(sheetData, sstPart, sharedCache, 3, 8, "Площ. м2", STYLE_CENTER10_BOLD_BORDER);

                        int row = 4;
                        foreach (var item in excelRows)
                        {
                            var ceilings = item.CeilingData.OrderBy(k => k.Key).ToList();
                            var walls = item.WallData.OrderBy(k => k.Key).ToList();
                            var cols = item.ColumnData.OrderBy(k => k.Key).ToList();
                            int rows = Math.Max(ceilings.Count, Math.Max(walls.Count, cols.Count));
                            if (rows == 0) continue;
                            int startRow = row;
                            int endRow = row + rows - 1;

                            for (int i = 0; i < rows; i++)
                            {
                                int curRow = row + i;
                                SetText(sheetData, sstPart, sharedCache, curRow, 1, item.RoomNumber, STYLE_CENTER10_BORDER);
                                SetText(sheetData, sstPart, sharedCache, curRow, 2, item.RoomName, STYLE_CENTER10_BORDER);
                                SetText(sheetData, sstPart, sharedCache, curRow, 3, i < ceilings.Count ? ceilings[i].Key : "", STYLE_LEFT10_BORDER);
                                if (i < ceilings.Count) SetNumber(sheetData, curRow, 4, ceilings[i].Value, STYLE_CENTER10_BORDER); else SetText(sheetData, sstPart, sharedCache, curRow, 4, "", STYLE_CENTER10_BORDER);
                                SetText(sheetData, sstPart, sharedCache, curRow, 5, i < walls.Count ? walls[i].Key : "", STYLE_LEFT10_BORDER);
                                if (i < walls.Count) SetNumber(sheetData, curRow, 6, walls[i].Value, STYLE_CENTER10_BORDER); else SetText(sheetData, sstPart, sharedCache, curRow, 6, "", STYLE_CENTER10_BORDER);
                                SetText(sheetData, sstPart, sharedCache, curRow, 7, i < cols.Count ? cols[i].Key : "", STYLE_LEFT10_BORDER);
                                if (i < cols.Count) SetNumber(sheetData, curRow, 8, cols[i].Value, STYLE_CENTER10_BORDER); else SetText(sheetData, sstPart, sharedCache, curRow, 8, "", STYLE_CENTER10_BORDER);
                                SetText(sheetData, sstPart, sharedCache, curRow, 9, "", STYLE_CENTER10_BORDER);
                            }
                            Merge(mergeCells, startRow, 1, endRow, 1); Merge(mergeCells, startRow, 2, endRow, 2); Merge(mergeCells, startRow, 9, endRow, 9);
                            row = endRow + 1;
                        }
                        sstPart.SharedStringTable.Save();
                        worksheetPart.Worksheet.Save();
                        workbookPart.Workbook.Save();
                    }
                }
                catch (Exception ex)
                {
                    roomBookToExcelProgressBarWPF.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.Close());
                    TaskDialog.Show("Revit", "Error: " + ex.Message);
                    return Result.Cancelled;
                }
            }
            return Result.Succeeded;
        }

        // ===================== OPENXML helpers (в этом же классе) =====================

        // Style indexes in Stylesheet.CellFormats:
        // 0 default
        // 1 Center + Wrap + ISOCPEUR 10
        // 2 Left   + Wrap + ISOCPEUR 10
        // 3 Center + Wrap + ISOCPEUR 14 Bold
        // 4 Center + Wrap + ISOCPEUR 10 + ThinBorder
        // 5 Left   + Wrap + ISOCPEUR 10 + ThinBorder
        // 6 Center + Wrap + ISOCPEUR 10 Bold
        // 7 Center + Wrap + ISOCPEUR 10 Bold + ThinBorder
        private const uint STYLE_CENTER10 = 1;
        private const uint STYLE_LEFT10 = 2;
        private const uint STYLE_CENTER14_BOLD = 3;
        private const uint STYLE_CENTER10_BORDER = 4;
        private const uint STYLE_LEFT10_BORDER = 5;
        private const uint STYLE_CENTER10_BOLD = 6;
        private const uint STYLE_CENTER10_BOLD_BORDER = 7;

        private static Stylesheet BuildStylesheet_ISOCPEUR()
        {
            // Fonts: 0 default, 1 ISO10, 2 ISO10 Bold, 3 ISO14 Bold
            Fonts fonts = new Fonts(
                new Font(),
                new Font(new FontName() { Val = "ISOCPEUR" }, new FontSize() { Val = 10 }),
                new Font(new Bold(), new FontName() { Val = "ISOCPEUR" }, new FontSize() { Val = 10 }),
                new Font(new Bold(), new FontName() { Val = "ISOCPEUR" }, new FontSize() { Val = 14 })
            );

            Fills fills = new Fills(
                new Fill(new PatternFill() { PatternType = PatternValues.None }),
                new Fill(new PatternFill() { PatternType = PatternValues.Gray125 })
            );

            Borders borders = new Borders(
                new Border(),
                new Border(
                    new LeftBorder() { Style = BorderStyleValues.Thin, Color = new Color() { Auto = true } },
                    new RightBorder() { Style = BorderStyleValues.Thin, Color = new Color() { Auto = true } },
                    new TopBorder() { Style = BorderStyleValues.Thin, Color = new Color() { Auto = true } },
                    new BottomBorder() { Style = BorderStyleValues.Thin, Color = new Color() { Auto = true } },
                    new DiagonalBorder()
                )
            );

            CellFormats cellFormats = new CellFormats(
                new CellFormat(), // 0 default

                // 1 center10
                new CellFormat
                {
                    FontId = 1,
                    FillId = 0,
                    BorderId = 0,
                    ApplyFont = true,
                    Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, WrapText = true }
                },

                // 2 left10
                new CellFormat
                {
                    FontId = 1,
                    FillId = 0,
                    BorderId = 0,
                    ApplyFont = true,
                    Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Center, WrapText = true }
                },

                // 3 center14 bold
                new CellFormat
                {
                    FontId = 3,
                    FillId = 0,
                    BorderId = 0,
                    ApplyFont = true,
                    Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, WrapText = true }
                },

                // 4 center10 + border
                new CellFormat
                {
                    FontId = 1,
                    FillId = 0,
                    BorderId = 1,
                    ApplyFont = true,
                    ApplyBorder = true,
                    Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, WrapText = true }
                },

                // 5 left10 + border
                new CellFormat
                {
                    FontId = 1,
                    FillId = 0,
                    BorderId = 1,
                    ApplyFont = true,
                    ApplyBorder = true,
                    Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Center, WrapText = true }
                },

                // 6 center10 bold
                new CellFormat
                {
                    FontId = 2,
                    FillId = 0,
                    BorderId = 0,
                    ApplyFont = true,
                    Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, WrapText = true }
                },

                // 7 center10 bold + border
                new CellFormat
                {
                    FontId = 2,
                    FillId = 0,
                    BorderId = 1,
                    ApplyFont = true,
                    ApplyBorder = true,
                    Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, WrapText = true }
                }
            );

            return new Stylesheet(fonts, fills, borders, cellFormats);
        }

        private static void AddColumn(Columns cols, uint min, uint max, double width)
        {
            cols.Append(new Column()
            {
                Min = min,
                Max = max,
                Width = width,
                CustomWidth = true
            });
        }

        private static void Merge(MergeCells mergeCells, int r1, int c1, int r2, int c2)
        {
            mergeCells.Append(new MergeCell()
            {
                Reference = new StringValue($"{GetColumnName(c1)}{r1}:{GetColumnName(c2)}{r2}")
            });
        }

        private static void SetText(
            SheetData sheetData,
            SharedStringTablePart sstPart,
            Dictionary<string, int> cache,
            int rowIndex,
            int colIndex,
            string text,
            uint styleIndex)
        {
            text ??= "";

            int sstIndex;
            if (!cache.TryGetValue(text, out sstIndex))
            {
                sstIndex = sstPart.SharedStringTable.ChildElements.Count;
                sstPart.SharedStringTable.AppendChild(new SharedStringItem(new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
                cache[text] = sstIndex;
            }

            Cell cell = InsertCell(sheetData, rowIndex, colIndex);
            cell.DataType = CellValues.SharedString;
            cell.CellValue = new CellValue(sstIndex.ToString(CultureInfo.InvariantCulture));
            cell.StyleIndex = styleIndex;
        }
        private static void SetNumber(SheetData sheetData, int rowIndex, int colIndex, double value, uint styleIndex)
        {
            Cell cell = InsertCell(sheetData, rowIndex, colIndex);
            cell.DataType = CellValues.Number;
            cell.CellValue = new CellValue(value.ToString("0.################", CultureInfo.InvariantCulture));
            cell.StyleIndex = styleIndex;
        }
        private static Cell InsertCell(SheetData sheetData, int rowIndex, int colIndex)
        {
            uint rIdx = (uint)rowIndex;
            string cellRef = GetColumnName(colIndex) + rowIndex.ToString(CultureInfo.InvariantCulture);

            Row row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex != null && r.RowIndex.Value == rIdx);
            if (row == null)
            {
                row = new Row() { RowIndex = rIdx };
                sheetData.Append(row);
            }

            Cell existing = row.Elements<Cell>().FirstOrDefault(c => c.CellReference != null && c.CellReference.Value == cellRef);
            if (existing != null) return existing;

            Cell refCell = null;
            foreach (Cell c in row.Elements<Cell>())
            {
                if (string.Compare(c.CellReference.Value, cellRef, StringComparison.Ordinal) > 0)
                {
                    refCell = c;
                    break;
                }
            }

            Cell cell = new Cell() { CellReference = cellRef };
            if (refCell != null) row.InsertBefore(cell, refCell);
            else row.Append(cell);

            return cell;
        }

        private static string GetColumnName(int colIndex1Based)
        {
            int dividend = colIndex1Based;
            string columnName = "";
            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }
            return columnName;
        }


        private static string GetStringParam(Element element, BuiltInParameter parameter)
        {
            return element?.get_Parameter(parameter)?.AsString() ?? string.Empty;
        }

        private static string GetStringParam(Element element, Guid parameterGuid)
        {
            return element?.get_Parameter(parameterGuid)?.AsString() ?? string.Empty;
        }

        private static string GetRoomName(Room room)
        {
            return GetStringParam(room, BuiltInParameter.ROOM_NAME);
        }

        private static string GetElementDescription(Element element, Guid parameterGuid)
        {
            string value = GetStringParam(element, parameterGuid);
            return string.IsNullOrWhiteSpace(value) ? "(без имени типа)" : value;
        }

        private void StartProgressWindow(int maximum)
        {
            using (ManualResetEventSlim readyEvent = new ManualResetEventSlim(false))
            {
                Thread newWindowThread = new Thread(() => ThreadStartingPoint(readyEvent));
                newWindowThread.SetApartmentState(ApartmentState.STA);
                newWindowThread.IsBackground = true;
                newWindowThread.Start();

                if (!readyEvent.Wait(TimeSpan.FromSeconds(5)))
                    throw new InvalidOperationException("Не удалось открыть окно прогресса.");
            }

            roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() =>
            {
                roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Minimum = 0;
                roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Maximum = maximum;
            });
        }

        private void AccumulateWallAreas(Document doc, IEnumerable<WallType> wallTypes, string roomNumber, Dictionary<string, double> itemData, Guid elemData, Guid roombookRoomNumber, bool isColumn, bool addPrefixForColumns = true)
        {
            if (wallTypes == null || itemData == null) return;

            foreach (WallType wallType in wallTypes)
            {
                string model = GetStringParam(wallType, BuiltInParameter.ALL_MODEL_MODEL);
                bool typeIsColumn = model.IndexOf("Колонны", StringComparison.OrdinalIgnoreCase) >= 0;
                if (isColumn && !typeIsColumn) continue;
                if (!isColumn && typeIsColumn) continue;

                var wallsOfType = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_Walls)
                    .OfClass(typeof(Wall))
                    .WhereElementIsNotElementType()
                    .Cast<Wall>()
                    .Where(w => w.WallType.Id == wallType.Id &&
                                w.get_Parameter(roombookRoomNumber) != null &&
                                w.get_Parameter(roombookRoomNumber).AsString() == roomNumber)
                    .ToList();

                double wallArea = 0;
                foreach (Wall wall in wallsOfType)
                {
#if R2019 || R2020 || R2021
                    wallArea += UnitUtils.ConvertFromInternalUnits(wall.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsDouble(), DisplayUnitType.DUT_SQUARE_METERS);
#else
                    wallArea += UnitUtils.ConvertFromInternalUnits(wall.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsDouble(), UnitTypeId.SquareMeters);
#endif
                }

                string name = GetElementDescription(wallType, elemData);
                string key = (isColumn && addPrefixForColumns) ? $"Колонны — {name}" : name;
                double value = Math.Round(wallArea, 2);
                if (itemData.ContainsKey(key)) itemData[key] += value;
                else itemData[key] = value;
            }
        }

        private void AccumulateCeilingAreas(Document doc, IEnumerable<CeilingType> ceilingTypes, string roomNumber, Dictionary<string, double> itemData, Guid elemData, Guid roombookRoomNumber)
        {
            if (ceilingTypes == null || itemData == null) return;

            foreach (CeilingType ceilingType in ceilingTypes)
            {
                if (ceilingType == null) continue;

                var ceilingsOfType = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_Ceilings)
                    .OfClass(typeof(Ceiling))
                    .WhereElementIsNotElementType()
                    .Cast<Ceiling>()
                    .Where(c => c.GetTypeId() == ceilingType.Id &&
                                c.get_Parameter(roombookRoomNumber) != null &&
                                c.get_Parameter(roombookRoomNumber).AsString() == roomNumber)
                    .ToList();

                double ceilingArea = 0;
                foreach (Ceiling ceiling in ceilingsOfType)
                {
#if R2019 || R2020 || R2021
                    ceilingArea += UnitUtils.ConvertFromInternalUnits(ceiling.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsDouble(), DisplayUnitType.DUT_SQUARE_METERS);
#else
                    ceilingArea += UnitUtils.ConvertFromInternalUnits(ceiling.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsDouble(), UnitTypeId.SquareMeters);
#endif
                }

                string name = GetStringParam(ceilingType, elemData);
                if (string.IsNullOrWhiteSpace(name)) name = ceilingType.Name ?? "(без имени типа)";
                double value = Math.Round(ceilingArea, 2);
                if (itemData.ContainsKey(name)) itemData[name] += value;
                else itemData[name] = value;
            }
        }
        // ===================== твои старые helpers ниже оставлены как есть =====================

        public class WallTypeIdComparer : IEqualityComparer<WallType>
        {
            public bool Equals(WallType x, WallType y) { return x.Id == y.Id; }
            public int GetHashCode(WallType obj) { return obj.Id.GetHashCode(); }
        }

        private void ThreadStartingPoint(ManualResetEventSlim readyEvent)
        {
            roomBookToExcelProgressBarWPF = new RoomBookToExcelProgressBarWPF();
            roomBookToExcelProgressBarWPF.Show();
            readyEvent.Set();
            System.Windows.Threading.Dispatcher.Run();
        }

        private static async Task GetPluginStartInfo()
        {
            Assembly thisAssembly = Assembly.GetExecutingAssembly();
            string assemblyName = "RoomBookToExcel";
            string assemblyNameRus = "RoomBook в Excel";
            string assemblyFolderPath = Path.GetDirectoryName(thisAssembly.Location);

            int lastBackslashIndex = assemblyFolderPath.LastIndexOf("\\");
            string dllPath = assemblyFolderPath.Substring(0, lastBackslashIndex + 1) + "PluginInfoCollector\\PluginInfoCollector.dll";

            Assembly assembly = Assembly.LoadFrom(dllPath);
            Type type = assembly.GetType("PluginInfoCollector.InfoCollector");

            if (type != null)
            {
                object instance = Activator.CreateInstance(type);
                var method = type.GetMethod("CollectPluginUsageAsync");
                if (method != null)
                {
                    Task task = (Task)method.Invoke(instance, new object[] { assemblyName, assemblyNameRus });
                    await task;
                }
            }
        }
    }
}

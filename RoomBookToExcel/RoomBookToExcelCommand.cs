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
            Guid roombookRoomName = new Guid("b59a22a9-7890-45bd-9f93-a186341eef58");
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
                Thread newWindowThread = new Thread(new ThreadStart(ThreadStartingPoint));
                newWindowThread.SetApartmentState(ApartmentState.STA);
                newWindowThread.IsBackground = true;
                newWindowThread.Start();

                int step = 0;
                Thread.Sleep(100);
                roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Minimum = 0);
                roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Maximum = roomList.Count);

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
                                    var m = f.FloorType.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString();
                                    return m == "Пол" || m == "Полы";
                                })
                                .Where(f => f.get_Parameter(roombookRoomNumber) != null)
                                .Where(f => f.get_Parameter(roombookRoomNumber).AsString() == room.Number)
                                .OrderBy(f => f.FloorType.get_Parameter(BuiltInParameter.WINDOW_TYPE_ID).AsString(), new AlphanumComparatorFastString())
                                .ToList();

                            // Стены
                            List<Wall> wallList = new FilteredElementCollector(doc)
                                .OfCategory(BuiltInCategory.OST_Walls)
                                .OfClass(typeof(Wall))
                                .WhereElementIsNotElementType()
                                .Cast<Wall>()
                                .Where(w => w.WallType.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL) != null &&
                                    !string.IsNullOrWhiteSpace(w.WallType.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString()))
                                .Where(w => w.WallType.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString() == "Отделка стен")
                                .Where(w => w.get_Parameter(roombookRoomNumber) != null)
                                .Where(w => w.get_Parameter(roombookRoomNumber).AsString() == room.Number)
                                .OrderBy(w => w.WallType.get_Parameter(BuiltInParameter.WINDOW_TYPE_ID).AsString(), new AlphanumComparatorFastString())
                                .ToList();

                            // Потолки
                            List<Ceiling> ceilingList = new FilteredElementCollector(doc)
                                .OfCategory(BuiltInCategory.OST_Ceilings)
                                .OfClass(typeof(Ceiling))
                                .WhereElementIsNotElementType()
                                .Cast<Ceiling>()
                                .Where(c => doc.GetElement(c.GetTypeId()).get_Parameter(BuiltInParameter.ALL_MODEL_MODEL) != null)
                                .Where(c =>
                                {
                                    var m = doc.GetElement(c.GetTypeId()).get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString();
                                    return m == "Потолок" || m == "Потолки";
                                })
                                .Where(w => w.get_Parameter(roombookRoomNumber) != null)
                                .Where(w => w.get_Parameter(roombookRoomNumber).AsString() == room.Number)
                                .OrderBy(c => doc.GetElement(c.GetTypeId()).get_Parameter(BuiltInParameter.WINDOW_TYPE_ID).AsString(), new AlphanumComparatorFastString())
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
                                .OrderBy(ft => ft.get_Parameter(BuiltInParameter.WINDOW_TYPE_ID).AsString(), new AlphanumComparatorFastString())
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
                                SetText(sheetData, sstPart, sharedCache, row, 2, room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString(), STYLE_CENTER10_BORDER);
                                SetText(sheetData, sstPart, sharedCache, row, 3, "Пол", STYLE_CENTER10_BORDER);
                                SetText(sheetData, sstPart, sharedCache, row, 4, floorType.get_Parameter(BuiltInParameter.WINDOW_TYPE_ID).AsString(), STYLE_CENTER10_BORDER);
                                SetText(sheetData, sstPart, sharedCache, row, 5, floorType.get_Parameter(elemData).AsString(), STYLE_LEFT10_BORDER);
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
                                .OrderBy(wt => wt.get_Parameter(BuiltInParameter.WINDOW_TYPE_ID).AsString(), new AlphanumComparatorFastString())
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
                                SetText(sheetData, sstPart, sharedCache, row, 2, room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString(), STYLE_CENTER10_BORDER);
                                SetText(sheetData, sstPart, sharedCache, row, 3, "Отделка\r\nстен", STYLE_CENTER10_BORDER);
                                SetText(sheetData, sstPart, sharedCache, row, 4, wallType.get_Parameter(BuiltInParameter.WINDOW_TYPE_ID).AsString(), STYLE_CENTER10_BORDER);
                                SetText(sheetData, sstPart, sharedCache, row, 5, wallType.get_Parameter(elemData).AsString(), STYLE_LEFT10_BORDER);
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
                                .OrderBy(ct => ct.get_Parameter(BuiltInParameter.WINDOW_TYPE_ID).AsString(), new AlphanumComparatorFastString())
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
                                SetText(sheetData, sstPart, sharedCache, row, 2, room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString(), STYLE_CENTER10_BORDER);
                                SetText(sheetData, sstPart, sharedCache, row, 3, "Потолок", STYLE_CENTER10_BORDER);
                                SetText(sheetData, sstPart, sharedCache, row, 4, ceilingType.get_Parameter(BuiltInParameter.WINDOW_TYPE_ID).AsString(), STYLE_CENTER10_BORDER);
                                SetText(sheetData, sstPart, sharedCache, row, 5, ceilingType.get_Parameter(elemData).AsString(), STYLE_LEFT10_BORDER);
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

            // Остальные ветки у тебя ниже пока Interop — перепишем по тому же шаблону
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
        // 6 Center + Wrap + ISOCPEUR 10 Bold + ThinBorder
        private const uint STYLE_CENTER10 = 1;
        private const uint STYLE_LEFT10 = 2;
        private const uint STYLE_CENTER14_BOLD = 3;
        private const uint STYLE_CENTER10_BORDER = 4;
        private const uint STYLE_LEFT10_BORDER = 5;
        private const uint STYLE_CENTER10_BOLD_BORDER = 6;

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
                // 6 center10 bold + border (для заголовков)
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
                sstIndex = sstPart.SharedStringTable.Count();
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

        // ===================== твои старые helpers ниже оставлены как есть =====================

        public class WallTypeIdComparer : IEqualityComparer<WallType>
        {
            public bool Equals(WallType x, WallType y) { return x.Id == y.Id; }
            public int GetHashCode(WallType obj) { return obj.Id.GetHashCode(); }
        }

        private void ThreadStartingPoint()
        {
            roomBookToExcelProgressBarWPF = new RoomBookToExcelProgressBarWPF();
            roomBookToExcelProgressBarWPF.Show();
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
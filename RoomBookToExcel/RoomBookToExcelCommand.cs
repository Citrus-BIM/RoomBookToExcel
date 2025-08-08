using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace RoomBookToExcel
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class RoomBookToExcelCommand : IExternalCommand
    {
        RoomBookToExcelProgressBarWPF roomBookToExcelProgressBarWPF;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                _ = GetPluginStartInfo();
            }
            catch { }

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
            if (roomBookToExcelWPF.DialogResult != true)
            {
                return Result.Cancelled;
            }
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
                    // Запускаем Excel
                    var excelApp = new Excel.Application();
                    excelApp.Visible = false;
                    excelApp.DisplayAlerts = false;

                    // Создаем книгу и лист
                    Excel.Workbook workbook = excelApp.Workbooks.Add();
                    Excel.Worksheet ws = workbook.Worksheets[1];
                    ws.Name = "RoomBook";

                    // Задаём ширину столбцов
                    ws.Columns[1].ColumnWidth = 10;
                    ws.Columns[2].ColumnWidth = 30;
                    ws.Columns[3].ColumnWidth = 10;
                    ws.Columns[4].ColumnWidth = 10;
                    ws.Columns[5].ColumnWidth = 50;
                    ws.Columns[6].ColumnWidth = 10;
                    ws.Columns[7].ColumnWidth = 15;
                    ws.Columns[8].ColumnWidth = 20;

                    // Заголовок 1 (объединение, стиль, шрифт)
                    Excel.Range rng1 = ws.Range[ws.Cells[1, 1], ws.Cells[1, 8]];
                    rng1.Merge();
                    rng1.Value2 = "Таблица вид 2";
                    rng1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    rng1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    rng1.WrapText = true;
                    rng1.Font.Name = "ISOCPEUR";
                    rng1.Font.Size = 10;

                    // Заголовок 2 (объединение, стиль, шрифт)
                    Excel.Range rng2 = ws.Range[ws.Cells[2, 1], ws.Cells[2, 8]];
                    rng2.Merge();
                    rng2.Value2 = "Румбук - Спецификация помещений";
                    rng2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    rng2.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    rng2.WrapText = true;
                    rng2.Font.Name = "ISOCPEUR";
                    rng2.Font.Size = 14;
                    rng2.Font.Bold = true;

                    // Заголовок 3 (объединение, стиль, шрифт)
                    Excel.Range rng3 = ws.Range[ws.Cells[3, 1], ws.Cells[3, 8]];
                    rng3.Merge();
                    rng3.Value2 = "Ссылки на листы документации";
                    rng3.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                    rng3.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    rng3.WrapText = true;
                    rng3.Font.Name = "ISOCPEUR";
                    rng3.Font.Size = 10;

                    // Заголовки столбцов
                    string[] headers = new string[]
                    {
                        "Номер помещения", "Имя помещения", "Тип элемента", "Марка элемента",
                        "Наименование элемента", "Ед. изм", "Кол-во", "Примечание"
                    };
                    for (int i = 0; i < headers.Length; i++)
                    {
                        ws.Cells[4, i + 1].Value2 = headers[i];
                    }

                    Excel.Range rngHeaders = ws.Range[ws.Cells[4, 1], ws.Cells[4, 8]];
                    rngHeaders.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    rngHeaders.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    rngHeaders.WrapText = true;
                    rngHeaders.Font.Name = "ISOCPEUR";
                    rngHeaders.Font.Size = 10;

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
                            .Where(f => {
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
                            .Where(w => w.WallType.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL) != null)
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

                        // Полы
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
                        floorTypesList = floorTypesList.OrderBy(ft => ft.get_Parameter(BuiltInParameter.WINDOW_TYPE_ID).AsString(), new AlphanumComparatorFastString()).ToList();
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
                            ws.Cells[row, 1].Value2 = room.Number;
                            ws.Cells[row, 2].Value2 = room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString();
                            ws.Cells[row, 3].Value2 = "Пол";
                            ws.Cells[row, 4].Value2 = floorType.get_Parameter(BuiltInParameter.WINDOW_TYPE_ID).AsString();
                            ws.Cells[row, 5].Value2 = floorType.get_Parameter(elemData).AsString();
                            ws.Cells[row, 6].Value2 = "м2";
                            ws.Cells[row, 7].Value2 = Math.Round(floorArea, 2);
                            row++;
                        }

                        // Стены
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
                        wallTypesList = wallTypesList.OrderBy(wt => wt.get_Parameter(BuiltInParameter.WINDOW_TYPE_ID).AsString(), new AlphanumComparatorFastString()).ToList();
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
                            ws.Cells[row, 1].Value2 = room.Number;
                            ws.Cells[row, 2].Value2 = room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString();
                            ws.Cells[row, 3].Value2 = "Отделка\r\nстен";
                            ws.Cells[row, 4].Value2 = wallType.get_Parameter(BuiltInParameter.WINDOW_TYPE_ID).AsString();
                            ws.Cells[row, 5].Value2 = wallType.get_Parameter(elemData).AsString();
                            ws.Cells[row, 6].Value2 = "м2";
                            ws.Cells[row, 7].Value2 = Math.Round(wallArea, 2);
                            row++;
                        }

                        // Потолки
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
                        ceilingTypesList = ceilingTypesList.OrderBy(ct => ct.get_Parameter(BuiltInParameter.WINDOW_TYPE_ID).AsString(), new AlphanumComparatorFastString()).ToList();
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
                            ws.Cells[row, 1].Value2 = room.Number;
                            ws.Cells[row, 2].Value2 = room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString();
                            ws.Cells[row, 3].Value2 = "Потолок";
                            ws.Cells[row, 4].Value2 = ceilingType.get_Parameter(BuiltInParameter.WINDOW_TYPE_ID).AsString();
                            ws.Cells[row, 5].Value2 = ceilingType.get_Parameter(elemData).AsString();
                            ws.Cells[row, 6].Value2 = "м2";
                            ws.Cells[row, 7].Value2 = Math.Round(ceilingArea, 2);
                            row++;
                        }

                        int endRow = row - 1;
                        ws.Range[ws.Cells[startRow, 1], ws.Cells[endRow, 1]].Merge();
                        ws.Range[ws.Cells[startRow, 2], ws.Cells[endRow, 2]].Merge();
                    }

                    roomBookToExcelProgressBarWPF.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.Close());

                    // Стилизация (шрифты, выравнивание, границы — если нужно, можно добавить по Range)
                    Excel.Range contentRange = ws.Range[ws.Cells[5, 1], ws.Cells[row - 1, 8]];
                    contentRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    contentRange.WrapText = true;
                    contentRange.Font.Name = "ISOCPEUR";
                    contentRange.Font.Size = 10;
                    ws.Range[ws.Cells[5, 1], ws.Cells[row - 1, 4]].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    ws.Range[ws.Cells[5, 5], ws.Cells[row - 1, 5]].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                    ws.Range[ws.Cells[5, 6], ws.Cells[row - 1, 8]].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                    // Границы — тонкие (можно добавить по вкусу)
                    Excel.Range borderRange = ws.Range[ws.Cells[4, 1], ws.Cells[row - 1, 8]];
                    borderRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    borderRange.Borders.Weight = Excel.XlBorderWeight.xlThin;

                    // Диалог сохранения
                    var saveDialog = new System.Windows.Forms.SaveFileDialog();
                    saveDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
                    var result = saveDialog.ShowDialog();
                    if (result == System.Windows.Forms.DialogResult.OK)
                    {
                        workbook.SaveAs(saveDialog.FileName);
                    }

                    // Освобождаем ресурсы
                    workbook.Close(false);
                    excelApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ws);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
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
                Thread newWindowThread = new Thread(new ThreadStart(ThreadStartingPoint));
                newWindowThread.SetApartmentState(ApartmentState.STA);
                newWindowThread.IsBackground = true;
                newWindowThread.Start();
                Thread.Sleep(100);
                int step = 0;
                roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Minimum = 0);
                roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Maximum = roomList.Count);

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
                        itemFloorFinishByRoom.RoomName = room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString();

                        // Полы в помещении
                        List<Floor> floorList = new FilteredElementCollector(doc)
                            .OfCategory(BuiltInCategory.OST_Floors)
                            .OfClass(typeof(Floor))
                            .WhereElementIsNotElementType()
                            .Cast<Floor>()
                            .Where(w => w.FloorType.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL) != null)
                            .Where(f => {
                                var model = f.FloorType.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString();
                                return model == "Пол" || model == "Полы";
                            })
                            .Where(w => w.get_Parameter(roombookRoomNumber) != null)
                            .Where(w => w.get_Parameter(roombookRoomNumber).AsString() == room.Number)
                            .OrderBy(w => w.FloorType.get_Parameter(BuiltInParameter.WINDOW_TYPE_ID).AsString(), new AlphanumComparatorFastString())
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
                            .OrderBy(wt => wt.get_Parameter(BuiltInParameter.WINDOW_TYPE_ID).AsString(), new AlphanumComparatorFastString())
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
                                        var model = f.FloorType.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString();
                                        return model == "Пол" || model == "Полы";
                                    })
                                    .Where(w => w.get_Parameter(roombookRoomNumber) != null)
                                    .Where(w => w.get_Parameter(roombookRoomNumber).AsString() == tmpItemFloorFinish.RoomNumber)
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

                                string floorTypeDisc = floorType.get_Parameter(elemData).AsString();
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

                    int itemDataCnt = itemFloorFinishByRoomExcelStringList.Max(i => i.ItemData.Count) * 2;

                    // --- EXCEL Interop ---
                    var excelApp = new Excel.Application();
                    excelApp.Visible = false;
                    excelApp.DisplayAlerts = false;
                    Excel.Workbook workbook = excelApp.Workbooks.Add();
                    Excel.Worksheet ws = workbook.Worksheets[1];
                    ws.Name = "FloorFinish";

                    // Задаем ширину столбцов
                    ws.Columns[1].ColumnWidth = 60;
                    ws.Columns[2].ColumnWidth = 120;
                    for (int i = 3; i <= itemDataCnt + 2; i += 2)
                    {
                        ws.Columns[i].ColumnWidth = 65;
                        ws.Columns[i + 1].ColumnWidth = 15;
                    }
                    ws.Columns[itemDataCnt + 3].ColumnWidth = 20;

                    // Заголовок 1 (строка 1)
                    Excel.Range rng1 = ws.Range[ws.Cells[1, 1], ws.Cells[1, itemDataCnt + 3]];
                    rng1.Merge();
                    rng1.Value2 = "Ведомость отделки стен";
                    rng1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    rng1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    rng1.WrapText = true;
                    rng1.Font.Name = "ISOCPEUR";
                    rng1.Font.Size = 10;
                    rng1.Font.Bold = true;

                    // Заголовок 2 (Номера помещений)
                    Excel.Range rngNum = ws.Range[ws.Cells[2, 1], ws.Cells[3, 1]];
                    rngNum.Merge();
                    rngNum.Value2 = "Номера помещений";
                    rngNum.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    rngNum.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    rngNum.WrapText = true;
                    rngNum.Font.Name = "ISOCPEUR";
                    rngNum.Font.Size = 10;

                    // Заголовок 3 (Наименования помещений)
                    Excel.Range rngName = ws.Range[ws.Cells[2, 2], ws.Cells[3, 2]];
                    rngName.Merge();
                    rngName.Value2 = "Наименования помещений";
                    rngName.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    rngName.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    rngName.WrapText = true;
                    rngName.Font.Name = "ISOCPEUR";
                    rngName.Font.Size = 10;

                    // Заголовок 4 (Типы отделки помещений)
                    Excel.Range rngType = ws.Range[ws.Cells[2, 3], ws.Cells[2, itemDataCnt + 2]];
                    rngType.Merge();
                    rngType.Value2 = "Типы отделки помещений";
                    rngType.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    rngType.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    rngType.WrapText = true;
                    rngType.Font.Name = "ISOCPEUR";
                    rngType.Font.Size = 10;

                    // Подзаголовки типов
                    int typeCnt = 1;
                    for (int i = 3; i <= itemDataCnt + 2; i += 2)
                    {
                        ws.Cells[3, i].Value2 = $"Отделка пола тип {typeCnt}";
                        ws.Cells[3, i + 1].Value2 = "Площ. м2";
                        typeCnt++;
                    }
                    Excel.Range rngTypeSub = ws.Range[ws.Cells[3, 3], ws.Cells[3, itemDataCnt + 2]];
                    rngTypeSub.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    rngTypeSub.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    rngTypeSub.WrapText = true;
                    rngTypeSub.Font.Name = "ISOCPEUR";
                    rngTypeSub.Font.Size = 10;

                    // Примечание
                    Excel.Range rngRemark = ws.Range[ws.Cells[2, itemDataCnt + 3], ws.Cells[3, itemDataCnt + 3]];
                    rngRemark.Merge();
                    rngRemark.Value2 = "Примечание";
                    rngRemark.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    rngRemark.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    rngRemark.WrapText = true;
                    rngRemark.Font.Name = "ISOCPEUR";
                    rngRemark.Font.Size = 10;

                    // Данные
                    for (int i = 0; i < itemFloorFinishByRoomExcelStringList.Count; i++)
                    {
                        ws.Cells[i + 4, 1].Value2 = itemFloorFinishByRoomExcelStringList[i].RoomNumber;
                        ws.Cells[i + 4, 2].Value2 = itemFloorFinishByRoomExcelStringList[i].RoomName;

                        for (int j = 0; j < itemFloorFinishByRoomExcelStringList[i].ItemData.Count; j++)
                        {
                            ws.Cells[i + 4, j * 2 + 3].Value2 = itemFloorFinishByRoomExcelStringList[i].ItemData.ElementAt(j).Key;
                            ws.Cells[i + 4, j * 2 + 3].WrapText = true;
                            ws.Cells[i + 4, j * 2 + 3].Font.Name = "ISOCPEUR";
                            ws.Cells[i + 4, j * 2 + 3].Font.Size = 10;
                            ws.Cells[i + 4, j * 2 + 3].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                            ws.Cells[i + 4, j * 2 + 3].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                            ws.Cells[i + 4, j * 2 + 4].Value2 = itemFloorFinishByRoomExcelStringList[i].ItemData.ElementAt(j).Value;
                            ws.Cells[i + 4, j * 2 + 4].WrapText = true;
                            ws.Cells[i + 4, j * 2 + 4].Font.Name = "ISOCPEUR";
                            ws.Cells[i + 4, j * 2 + 4].Font.Size = 10;
                            ws.Cells[i + 4, j * 2 + 4].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                            ws.Cells[i + 4, j * 2 + 4].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                        }
                    }

                    // Форматирование (левая колонка)
                    Excel.Range rngFormat = ws.Range[ws.Cells[4, 1], ws.Cells[itemFloorFinishByRoomExcelStringList.Count + 3, 2]];
                    rngFormat.WrapText = true;
                    rngFormat.Font.Name = "ISOCPEUR";
                    rngFormat.Font.Size = 10;
                    rngFormat.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    rngFormat.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                    // Рамки
                    Excel.Range rngBorder = ws.Range[ws.Cells[2, 1], ws.Cells[itemFloorFinishByRoomExcelStringList.Count + 3, itemDataCnt + 3]];
                    rngBorder.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    rngBorder.Borders.Weight = Excel.XlBorderWeight.xlThin;

                    // Диалог сохранения
                    var saveDialog = new System.Windows.Forms.SaveFileDialog();
                    saveDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
                    var result = saveDialog.ShowDialog();
                    if (result == System.Windows.Forms.DialogResult.OK)
                    {
                        workbook.SaveAs(saveDialog.FileName);
                    }

                    workbook.Close(false);
                    excelApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ws);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
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
                // ---------- 0. Прогресс-окно ----------
                var uiThread = new Thread(new ThreadStart(ThreadStartingPoint)) { IsBackground = true };
                uiThread.SetApartmentState(ApartmentState.STA);
                uiThread.Start();
                Thread.Sleep(100);

                roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() =>
                {
                    roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Minimum = 0;
                    roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Maximum = roomList.Count;
                });

                try
                {
                    // ---------- 1. Сбор данных ----------
                    int step = 0;
                    var combos = new List<ItemWallFinishByRoom>();

                    foreach (Room room in roomList)
                    {
                        step++;
                        roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() =>
                        {
                            roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Value = step;
                            roomBookToExcelProgressBarWPF.label_ItemName.Content =
                                $"Сбор данных. Шаг {step} из {roomList.Count}";
                        });

                        var item = new ItemWallFinishByRoom
                        {
                            RoomNumber = room.Number,
                            RoomName = room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString()
                        };

                        var finishes = new FilteredElementCollector(doc)
                            .OfCategory(BuiltInCategory.OST_Walls)
                            .OfClass(typeof(Wall))
                            .WhereElementIsNotElementType()
                            .Cast<Wall>()
                            .Where(w =>
                                w.WallType.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL) != null &&
                                w.WallType.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString().StartsWith("Отделка стен") &&
                                w.get_Parameter(roombookRoomNumber) != null &&
                                w.get_Parameter(roombookRoomNumber).AsString() == room.Number)
                            .ToList();

                        var wallTypes = finishes
                            .Where(w => !w.WallType.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL)
                                                    .AsString().Contains("Колонны"))
                            .Select(w => w.WallType)
                            .Distinct(new WallTypeIdComparer())
                            .OrderBy(t => t.Name, StringComparer.Ordinal)
                            .ToList();

                        var columnTypes = finishes
                            .Where(w => w.WallType.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL)
                                                    .AsString().Contains("Колонны"))
                            .Select(w => w.WallType)
                            .Distinct(new WallTypeIdComparer())
                            .OrderBy(t => t.Name, StringComparer.Ordinal)
                            .ToList();

                        item.WallFinishes.AddRange(wallTypes);
                        item.ColumnFinishes.AddRange(columnTypes);
                        combos.Add(item);
                    }

                    // ---------- 2. Уникальные сочетания ----------
                    var uniqueCombos = combos
                        .Distinct(new ItemWallFinishByRoomComparer())
                        .OrderBy(c => c.RoomNumber, new AlphanumComparatorFastString())
                        .ToList();

                    roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() =>
                        roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Maximum = uniqueCombos.Count);

                    var excelRows = new List<ItemWallFinishByRoomExcelString>(); step = 0;

                    foreach (var combo in uniqueCombos)
                    {
                        step++;
                        roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() =>
                        {
                            roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Value = step;
                            roomBookToExcelProgressBarWPF.label_ItemName.Content =
                                $"Комбинация {step} из {uniqueCombos.Count}";
                        });

                        var row = new ItemWallFinishByRoomExcelString();
                        var rooms = combos.Where(c => c.Equals(combo)).ToList();

                        row.RoomNumber = string.Join(", ",
                            rooms.Select(r => r.RoomNumber).Distinct()
                                 .OrderBy(s => s, new AlphanumComparatorFastString()));

                        row.RoomName = string.Join(", ",
                            rooms.Select(r => r.RoomName).Distinct(StringComparer.Ordinal)
                                 .OrderBy(s => s, new AlphanumComparatorFastString()));

                        foreach (var r in rooms)
                        {
                            AccumulateArea(doc, r.WallFinishes, r.RoomNumber, row.ItemData, elemData, roombookRoomNumber, false);
                            AccumulateArea(doc, r.ColumnFinishes, r.RoomNumber, row.ItemData, elemData, roombookRoomNumber, true);
                        }

                        if (row.ItemData.Count > 0)
                            excelRows.Add(row);
                    }

                    roomBookToExcelProgressBarWPF.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.Close());

                    // ---------- 3. Excel (вертикальный вывод) ----------
                    var xl = new Excel.Application { Visible = false, DisplayAlerts = false };
                    var wb = xl.Workbooks.Add();
                    var ws = wb.Worksheets[1]; ws.Name = "Отделка стен";

                    ws.Columns[1].ColumnWidth = 60;
                    ws.Columns[2].ColumnWidth = 60;
                    ws.Columns[3].ColumnWidth = 15;
                    ws.Columns[4].ColumnWidth = 60;
                    ws.Columns[5].ColumnWidth = 15;
                    ws.Columns[6].ColumnWidth = 60;
                    ws.Columns[7].ColumnWidth = 15;
                    ws.Columns[8].ColumnWidth = 16.5;

                    FormatHeader8cols_BlockCeiling(ws);

                    const string CEILING_NOTE = "см. ведомость отделки потолков на листе";

                    int curRow = 4;
                    foreach (var r in excelRows)
                    {
                        var walls = r.ItemData
                            .Where(k => !k.Key.StartsWith("Колонны — "))
                            .OrderBy(k => k.Key)
                            .ToList();

                        var cols = r.ItemData
                            .Where(k => k.Key.StartsWith("Колонны — "))
                            .OrderBy(k => k.Key)
                            .Select(kv => new KeyValuePair<string, double>(
                                kv.Key.Replace("Колонны — ", ""), kv.Value)) // <-- без префикса в Excel
                            .ToList();

                        int rows = Math.Max(walls.Count, cols.Count);
                        if (rows == 0) continue;

                        int blockEnd = curRow + rows - 1;

                        // --- № и Имя ---
                        var rngNum = ws.Range[ws.Cells[curRow, 1], ws.Cells[blockEnd, 1]];
                        var rngName = ws.Range[ws.Cells[curRow, 2], ws.Cells[blockEnd, 2]];
                        rngNum.Merge(); rngName.Merge();
                        rngNum.Value2 = r.RoomNumber;
                        rngName.Value2 = r.RoomName;
                        foreach (var rng in new[] { rngNum, rngName })
                        {
                            rng.WrapText = true;
                            rng.ShrinkToFit = true;
                            rng.Font.Name = "ISOCPEUR";
                            rng.Font.Size = 10;
                            rng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                            rng.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                        }

                        // --- Отделка потолка (C) по блоку ---
                        var rngCeil = ws.Range[ws.Cells[curRow, 3], ws.Cells[blockEnd, 3]];
                        rngCeil.Merge();
                        rngCeil.Value2 = CEILING_NOTE;
                        rngCeil.WrapText = true;
                        rngCeil.ShrinkToFit = true;
                        rngCeil.Font.Name = "ISOCPEUR";
                        rngCeil.Font.Size = 10;
                        rngCeil.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        rngCeil.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                        // --- Строки данных ---
                        for (int i = 0; i < rows; i++)
                        {
                            int row = curRow + i;
                            if (i < walls.Count)
                            {
                                ws.Cells[row, 4].Value2 = walls[i].Key;
                                ws.Cells[row, 5].Value2 = walls[i].Value;
                            }
                            if (i < cols.Count)
                            {
                                ws.Cells[row, 6].Value2 = cols[i].Key;
                                ws.Cells[row, 7].Value2 = cols[i].Value;
                            }
                            StyleDataRow_BlockCeiling(ws, row);
                        }

                        ws.Range[ws.Cells[curRow, 1], ws.Cells[blockEnd, 8]].EntireRow.AutoFit();
                        curRow = blockEnd + 1;
                    }

                    var border = ws.Range[ws.Cells[2, 1], ws.Cells[curRow - 1, 8]];
                    border.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    border.Borders.Weight = Excel.XlBorderWeight.xlThin;

                    var used = ws.Range[ws.Cells[1, 1], ws.Cells[curRow - 1, 8]];
                    used.WrapText = true;
                    used.Rows.AutoFit();

                    var dlg = new System.Windows.Forms.SaveFileDialog { Filter = "Excel files (*.xlsx)|*.xlsx" };
                    if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                        wb.SaveAs(dlg.FileName);

                    wb.Close(false);
                    xl.Quit();
                    GC.Collect();
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
                Thread newWindowThread = new Thread(new ThreadStart(ThreadStartingPoint));
                newWindowThread.SetApartmentState(ApartmentState.STA);
                newWindowThread.IsBackground = true;
                newWindowThread.Start();
                Thread.Sleep(100);
                int step = 0;
                roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Minimum = 0);
                roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Maximum = roomList.Count);

                try
                {
                    List<ItemCeilingFinishByRoom> itemCeilingFinishByRoomList = new List<ItemCeilingFinishByRoom>();
                    foreach (Room room in roomList)
                    {
                        step++;
                        roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() =>
                            roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Value = step);
                        roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() =>
                            roomBookToExcelProgressBarWPF.label_ItemName.Content =
                                $"Сбор данных об отделке потолка. Шаг {step} из {roomList.Count}");

                        ItemCeilingFinishByRoom itemCeilingFinishByRoom = new ItemCeilingFinishByRoom();
                        itemCeilingFinishByRoom.RoomNumber = room.Number;
                        itemCeilingFinishByRoom.RoomName = room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString();

                        // Потолки в помещении
                        List<Ceiling> ceilingList = new FilteredElementCollector(doc)
                            .OfCategory(BuiltInCategory.OST_Ceilings)
                            .OfClass(typeof(Ceiling))
                            .WhereElementIsNotElementType()
                            .Cast<Ceiling>()
                            .Where(c => doc.GetElement(c.GetTypeId()).get_Parameter(BuiltInParameter.ALL_MODEL_MODEL) != null)
                            .Where(c =>
                                doc.GetElement(c.GetTypeId()).get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString() == "Потолок" ||
                                doc.GetElement(c.GetTypeId()).get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString() == "Потолки")
                            .Where(c => c.get_Parameter(roombookRoomNumber) != null)
                            .Where(c => c.get_Parameter(roombookRoomNumber).AsString() == room.Number)
                            .OrderBy(c => doc.GetElement(c.GetTypeId()).get_Parameter(BuiltInParameter.WINDOW_TYPE_ID).AsString(), new AlphanumComparatorFastString())
                            .ToList();

                        // Список уникальных типов потолков
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

                        itemCeilingFinishByRoom.CeilingTypesList = ceilingTypesList
                            .OrderBy(wt => wt.get_Parameter(BuiltInParameter.WINDOW_TYPE_ID).AsString(), new AlphanumComparatorFastString())
                            .ToList();

                        itemCeilingFinishByRoomList.Add(itemCeilingFinishByRoom);
                    }

                    List<ItemCeilingFinishByRoom> uniqueCeilingFinishSet =
                        itemCeilingFinishByRoomList.Distinct(new ItemCeilingFinishByRoomComparer()).ToList();

                    step = 0;
                    roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() =>
                        roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Minimum = 0);
                    roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() =>
                        roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Maximum = uniqueCeilingFinishSet.Count);

                    List<ItemCeilingFinishByRoomExcelString> itemCeilingFinishByRoomExcelStringList =
                        new List<ItemCeilingFinishByRoomExcelString>();

                    foreach (ItemCeilingFinishByRoom uniqueCeilingFinish in uniqueCeilingFinishSet)
                    {
                        step++;
                        roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() =>
                            roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Value = step);
                        roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() =>
                            roomBookToExcelProgressBarWPF.label_ItemName.Content =
                                $"Обработка сочетаний отделок. Шаг {step} из {uniqueCeilingFinishSet.Count}");

                        ItemCeilingFinishByRoomExcelString itemCeilingFinishByRoomExcelString = new ItemCeilingFinishByRoomExcelString();
                        itemCeilingFinishByRoomExcelString.ItemData = new Dictionary<string, double>();
                        List<ItemCeilingFinishByRoom> tmpItemCeilingFinishList = itemCeilingFinishByRoomList
                            .Where(i => i.Equals(uniqueCeilingFinish))
                            .OrderBy(i => i.RoomNumber, new AlphanumComparatorFastString())
                            .ToList();

                        List<string> roomNumbersList = new List<string>();
                        List<string> roomNamesList = new List<string>();

                        foreach (ItemCeilingFinishByRoom tmpItemCeilingFinish in tmpItemCeilingFinishList)
                        {
                            if (!roomNumbersList.Contains(tmpItemCeilingFinish.RoomNumber))
                                roomNumbersList.Add(tmpItemCeilingFinish.RoomNumber);

                            if (!roomNamesList.Contains(tmpItemCeilingFinish.RoomName))
                                roomNamesList.Add(tmpItemCeilingFinish.RoomName);

                            foreach (CeilingType ceilingType in tmpItemCeilingFinish.CeilingTypesList)
                            {
                                List<Ceiling> tmpCeilingList = new FilteredElementCollector(doc)
                                    .OfCategory(BuiltInCategory.OST_Ceilings)
                                    .OfClass(typeof(Ceiling))
                                    .WhereElementIsNotElementType()
                                    .Cast<Ceiling>()
                                    .Where(c => doc.GetElement(c.GetTypeId()).get_Parameter(BuiltInParameter.ALL_MODEL_MODEL) != null)
                                    .Where(c =>
                                        doc.GetElement(c.GetTypeId()).get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString() == "Потолок" ||
                                        doc.GetElement(c.GetTypeId()).get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString() == "Потолки")
                                    .Where(c => c.get_Parameter(roombookRoomNumber) != null)
                                    .Where(c => c.get_Parameter(roombookRoomNumber).AsString() == tmpItemCeilingFinish.RoomNumber)
                                    .Where(c => doc.GetElement(c.GetTypeId()).Id == ceilingType.Id)
                                    .ToList();

                                double ceilingArea = 0;
                                foreach (Ceiling ceiling in tmpCeilingList)
                                {
#if R2019 || R2020 || R2021
                                    ceilingArea += UnitUtils.ConvertFromInternalUnits(ceiling.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsDouble(), DisplayUnitType.DUT_SQUARE_METERS);
#else
                                    ceilingArea += UnitUtils.ConvertFromInternalUnits(ceiling.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsDouble(), UnitTypeId.SquareMeters);
#endif
                                }

                                string ceilingTypeDisc = ceilingType.get_Parameter(elemData).AsString();
                                if (itemCeilingFinishByRoomExcelString.ItemData.ContainsKey(ceilingTypeDisc))
                                {
                                    itemCeilingFinishByRoomExcelString.ItemData[ceilingTypeDisc] += Math.Round(ceilingArea, 2);
                                }
                                else
                                {
                                    itemCeilingFinishByRoomExcelString.ItemData.Add(ceilingTypeDisc, Math.Round(ceilingArea, 2));
                                }
                            }
                        }

                        itemCeilingFinishByRoomExcelString.RoomNumber = string.Join(", ", roomNumbersList.OrderBy(n => n, new AlphanumComparatorFastString()));
                        itemCeilingFinishByRoomExcelString.RoomName = string.Join(", ", roomNamesList.OrderBy(n => n, new AlphanumComparatorFastString()));
                        itemCeilingFinishByRoomExcelStringList.Add(itemCeilingFinishByRoomExcelString);
                    }
                    roomBookToExcelProgressBarWPF.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.Close());

                    int itemDataCnt = itemCeilingFinishByRoomExcelStringList.Max(i => i.ItemData.Count) * 2;

                    // ---- Excel Interop ----
                    var excelApp = new Excel.Application();
                    excelApp.Visible = false;
                    excelApp.DisplayAlerts = false;
                    Excel.Workbook workbook = excelApp.Workbooks.Add();
                    Excel.Worksheet ws = workbook.Worksheets[1];
                    ws.Name = "CeilingFinish";

                    ws.Columns[1].ColumnWidth = 60;
                    ws.Columns[2].ColumnWidth = 120;
                    for (int i = 3; i <= itemDataCnt + 2; i += 2)
                    {
                        ws.Columns[i].ColumnWidth = 65;
                        ws.Columns[i + 1].ColumnWidth = 15;
                    }
                    ws.Columns[itemDataCnt + 3].ColumnWidth = 20;

                    // Заголовок 1 (строка 1)
                    Excel.Range rng1 = ws.Range[ws.Cells[1, 1], ws.Cells[1, itemDataCnt + 3]];
                    rng1.Merge();
                    rng1.Value2 = "Ведомость отделки потолка";
                    rng1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    rng1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    rng1.WrapText = true;
                    rng1.Font.Name = "ISOCPEUR";
                    rng1.Font.Size = 10;
                    rng1.Font.Bold = true;

                    // Заголовок 2 (Номера помещений)
                    Excel.Range rngNum = ws.Range[ws.Cells[2, 1], ws.Cells[3, 1]];
                    rngNum.Merge();
                    rngNum.Value2 = "Номера помещений";
                    rngNum.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    rngNum.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    rngNum.WrapText = true;
                    rngNum.Font.Name = "ISOCPEUR";
                    rngNum.Font.Size = 10;

                    // Заголовок 3 (Наименования помещений)
                    Excel.Range rngName = ws.Range[ws.Cells[2, 2], ws.Cells[3, 2]];
                    rngName.Merge();
                    rngName.Value2 = "Наименования помещений";
                    rngName.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    rngName.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    rngName.WrapText = true;
                    rngName.Font.Name = "ISOCPEUR";
                    rngName.Font.Size = 10;

                    // Заголовок 4 (Типы отделки помещений)
                    Excel.Range rngType = ws.Range[ws.Cells[2, 3], ws.Cells[2, itemDataCnt + 2]];
                    rngType.Merge();
                    rngType.Value2 = "Типы отделки помещений";
                    rngType.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    rngType.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    rngType.WrapText = true;
                    rngType.Font.Name = "ISOCPEUR";
                    rngType.Font.Size = 10;

                    // Подзаголовки типов
                    int typeCnt = 1;
                    for (int i = 3; i <= itemDataCnt + 2; i += 2)
                    {
                        ws.Cells[3, i].Value2 = $"Отделка потолка тип {typeCnt}";
                        ws.Cells[3, i + 1].Value2 = "Площ. м2";
                        typeCnt++;
                    }
                    Excel.Range rngTypeSub = ws.Range[ws.Cells[3, 3], ws.Cells[3, itemDataCnt + 2]];
                    rngTypeSub.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    rngTypeSub.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    rngTypeSub.WrapText = true;
                    rngTypeSub.Font.Name = "ISOCPEUR";
                    rngTypeSub.Font.Size = 10;

                    // Примечание
                    Excel.Range rngRemark = ws.Range[ws.Cells[2, itemDataCnt + 3], ws.Cells[3, itemDataCnt + 3]];
                    rngRemark.Merge();
                    rngRemark.Value2 = "Примечание";
                    rngRemark.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    rngRemark.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    rngRemark.WrapText = true;
                    rngRemark.Font.Name = "ISOCPEUR";
                    rngRemark.Font.Size = 10;

                    // Данные
                    for (int i = 0; i < itemCeilingFinishByRoomExcelStringList.Count; i++)
                    {
                        ws.Cells[i + 4, 1].Value2 = itemCeilingFinishByRoomExcelStringList[i].RoomNumber;
                        ws.Cells[i + 4, 2].Value2 = itemCeilingFinishByRoomExcelStringList[i].RoomName;

                        for (int j = 0; j < itemCeilingFinishByRoomExcelStringList[i].ItemData.Count; j++)
                        {
                            ws.Cells[i + 4, j * 2 + 3].Value2 = itemCeilingFinishByRoomExcelStringList[i].ItemData.ElementAt(j).Key;
                            ws.Cells[i + 4, j * 2 + 3].WrapText = true;
                            ws.Cells[i + 4, j * 2 + 3].Font.Name = "ISOCPEUR";
                            ws.Cells[i + 4, j * 2 + 3].Font.Size = 10;
                            ws.Cells[i + 4, j * 2 + 3].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                            ws.Cells[i + 4, j * 2 + 3].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                            ws.Cells[i + 4, j * 2 + 4].Value2 = itemCeilingFinishByRoomExcelStringList[i].ItemData.ElementAt(j).Value;
                            ws.Cells[i + 4, j * 2 + 4].WrapText = true;
                            ws.Cells[i + 4, j * 2 + 4].Font.Name = "ISOCPEUR";
                            ws.Cells[i + 4, j * 2 + 4].Font.Size = 10;
                            ws.Cells[i + 4, j * 2 + 4].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                            ws.Cells[i + 4, j * 2 + 4].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                        }
                    }

                    // Форматирование
                    Excel.Range rngFormat = ws.Range[ws.Cells[4, 1], ws.Cells[itemCeilingFinishByRoomExcelStringList.Count + 3, 2]];
                    rngFormat.WrapText = true;
                    rngFormat.Font.Name = "ISOCPEUR";
                    rngFormat.Font.Size = 10;
                    rngFormat.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    rngFormat.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                    // Рамки
                    Excel.Range rngBorder = ws.Range[ws.Cells[2, 1], ws.Cells[itemCeilingFinishByRoomExcelStringList.Count + 3, itemDataCnt + 3]];
                    rngBorder.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    rngBorder.Borders.Weight = Excel.XlBorderWeight.xlThin;

                    // Диалог сохранения
                    var saveDialog = new System.Windows.Forms.SaveFileDialog();
                    saveDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
                    var result = saveDialog.ShowDialog();
                    if (result == System.Windows.Forms.DialogResult.OK)
                    {
                        workbook.SaveAs(saveDialog.FileName);
                    }

                    workbook.Close(false);
                    excelApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ws);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                }
                catch (Exception ex)
                {
                    roomBookToExcelProgressBarWPF.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.Close());
                    string errorMessage = "Error: " + ex.Message + " Line: " + ex.Source;
                    TaskDialog.Show("Revit", errorMessage);
                    return Result.Cancelled;
                }
            }
            else if (exportOptionName == "rbt_WallFinishByCombinationWithCeiling")
            {
                // ---------- 0. прогресс ----------
                var uiThread = new Thread(new ThreadStart(ThreadStartingPoint)) { IsBackground = true };
                uiThread.SetApartmentState(ApartmentState.STA);
                uiThread.Start();
                Thread.Sleep(100);

                roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() =>
                {
                    roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Minimum = 0;
                    roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Maximum = roomList.Count;
                });

                try
                {
                    // ---------- 1. сбор данных ----------
                    int step = 0;
                    var combos = new List<ItemWallFinishByRoomWithCeil>();

                    foreach (Room room in roomList)
                    {
                        step++;
                        roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() =>
                        {
                            roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Value = step;
                            roomBookToExcelProgressBarWPF.label_ItemName.Content = $"Сбор данных. Шаг {step} из {roomList.Count}";
                        });

                        var item = new ItemWallFinishByRoomWithCeil
                        {
                            RoomNumber = room.Number,
                            RoomName = room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString()
                        };

                        // стены/колонны
                        var finishes = new FilteredElementCollector(doc)
                            .OfCategory(BuiltInCategory.OST_Walls)
                            .OfClass(typeof(Wall))
                            .WhereElementIsNotElementType()
                            .Cast<Wall>()
                            .Where(w =>
                                w.WallType.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL) != null &&
                                w.WallType.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString().StartsWith("Отделка стен") &&
                                w.get_Parameter(roombookRoomNumber) != null &&
                                w.get_Parameter(roombookRoomNumber).AsString() == room.Number)
                            .ToList();

                        var wallTypes = finishes
                            .Where(w => !w.WallType.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString().Contains("Колонны"))
                            .Select(w => w.WallType)
                            .Distinct(new WallTypeIdComparer())
                            .OrderBy(t => t.Name, StringComparer.Ordinal)
                            .ToList();

                        var columnTypes = finishes
                            .Where(w => w.WallType.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString().Contains("Колонны"))
                            .Select(w => w.WallType)
                            .Distinct(new WallTypeIdComparer())
                            .OrderBy(t => t.Name, StringComparer.Ordinal)
                            .ToList();

                        item.WallFinishes.AddRange(wallTypes);
                        item.ColumnFinishes.AddRange(columnTypes);

                        // потолки — список типов
                        var ceilingTypes = new FilteredElementCollector(doc)
                            .OfCategory(BuiltInCategory.OST_Ceilings)
                            .OfClass(typeof(Ceiling))
                            .WhereElementIsNotElementType()
                            .Cast<Ceiling>()
                            .Where(c => c.get_Parameter(roombookRoomNumber) != null &&
                                        c.get_Parameter(roombookRoomNumber).AsString() == room.Number)
                            .Select(c => (CeilingType)doc.GetElement(c.GetTypeId()))
                            .Where(ct => ct != null &&
                                         ct.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL) != null &&
                                        (ct.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString() == "Потолок" ||
                                         ct.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString() == "Потолки"))
                            .Distinct(new ElementTypeIdComparer<CeilingType>())
                            .OrderBy(ct => ct.get_Parameter(BuiltInParameter.WINDOW_TYPE_ID)?.AsString() ?? ct.Name,
                                     new AlphanumComparatorFastString())
                            .ToList();

                        item.CeilingFinishes.AddRange(ceilingTypes);

                        combos.Add(item);
                    }

                    // ---------- 2. уникальные сочетания ----------
                    var uniqueCombos = combos
                        .Distinct(new ItemWallFinishByRoomWithCeilComparer())
                        .OrderBy(c => c.RoomNumber, new AlphanumComparatorFastString())
                        .ToList();

                    roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() =>
                        roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Maximum = uniqueCombos.Count);

                    var excelRows = new List<ItemWallFinishByRoomWithCeilExcelRow>();
                    step = 0;

                    foreach (var combo in uniqueCombos)
                    {
                        step++;
                        roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() =>
                        {
                            roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Value = step;
                            roomBookToExcelProgressBarWPF.label_ItemName.Content = $"Комбинация {step} из {uniqueCombos.Count}";
                        });

                        var row = new ItemWallFinishByRoomWithCeilExcelRow();
                        var roomsOfCombo = combos.Where(c => c.Equals(combo)).ToList();

                        row.RoomNumber = string.Join(", ",
                            roomsOfCombo.Select(r => r.RoomNumber).Distinct()
                                        .OrderBy(s => s, new AlphanumComparatorFastString()));
                        row.RoomName = string.Join(", ",
                            roomsOfCombo.Select(r => r.RoomName).Distinct(StringComparer.Ordinal)
                                        .OrderBy(s => s, new AlphanumComparatorFastString()));

                        // стены/колонны
                        foreach (var r in roomsOfCombo)
                        {
                            AccumulateArea(doc, r.WallFinishes, r.RoomNumber, row.WallData, elemData, roombookRoomNumber, false, addPrefixForColumns: true);
                            AccumulateArea(doc, r.ColumnFinishes, r.RoomNumber, row.ColumnData, elemData, roombookRoomNumber, true, addPrefixForColumns: true);
                        }

                        // потолки — считаем по типам, входящим в текущее сочетание
                        foreach (var r in roomsOfCombo)
                        {
                            AccumulateCeilingArea(doc, combo.CeilingFinishes, r.RoomNumber,
                                                  row.CeilingData, elemData, roombookRoomNumber);
                        }

                        if (row.WallData.Count > 0 || row.ColumnData.Count > 0 || row.CeilingData.Count > 0)
                            excelRows.Add(row);
                    }

                    roomBookToExcelProgressBarWPF.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.Close());

                    // ---------- 3. Excel ----------
                    var xl = new Excel.Application { Visible = false, DisplayAlerts = false };
                    var wb = xl.Workbooks.Add();
                    var ws = wb.Worksheets[1]; ws.Name = "Стены и Потолки";

                    // ширины (добавили I: Примечания)
                    ws.Columns[1].ColumnWidth = 50;   // A: №
                    ws.Columns[2].ColumnWidth = 50;   // B: Имя
                    ws.Columns[3].ColumnWidth = 45;   // C: Потолок тип
                    ws.Columns[4].ColumnWidth = 15;   // D: Потолок, м²
                    ws.Columns[5].ColumnWidth = 45;   // E: Стены тип
                    ws.Columns[6].ColumnWidth = 15;   // F: Стены, м²
                    ws.Columns[7].ColumnWidth = 50;   // G: Колонны тип
                    ws.Columns[8].ColumnWidth = 15;   // H: Колонны, м²
                    ws.Columns[9].ColumnWidth = 16; // I: Примечания

                    // Шапка: C–H = три пары + I «Примечания»
                    FormatHeader9cols_TriplePairs(ws);

                    int curRow = 4;
                    foreach (var r in excelRows)
                    {
                        var ceil = r.CeilingData.OrderBy(k => k.Key).ToList();
                        var walls = r.WallData.OrderBy(k => k.Key).ToList();
                        // для вывода колонн убираем префикс "Колонны — "
                        var cols = r.ColumnData.OrderBy(k => k.Key)
                                                 .Select(kv => new KeyValuePair<string, double>(
                                                     kv.Key.StartsWith("Колонны — ", StringComparison.Ordinal)
                                                         ? kv.Key.Substring("Колонны — ".Length)
                                                         : kv.Key,
                                                     kv.Value))
                                                 .ToList();

                        int rows = Math.Max(ceil.Count, Math.Max(walls.Count, cols.Count));
                        if (rows == 0) continue;

                        int blockEnd = curRow + rows - 1;

                        // A–B объединяем
                        var rngNum = ws.Range[ws.Cells[curRow, 1], ws.Cells[blockEnd, 1]];
                        var rngName = ws.Range[ws.Cells[curRow, 2], ws.Cells[blockEnd, 2]];
                        rngNum.Merge(); rngNum.Value2 = r.RoomNumber;
                        rngName.Merge(); rngName.Value2 = r.RoomName;
                        foreach (var rng in new[] { rngNum, rngName })
                        {
                            rng.WrapText = true; rng.ShrinkToFit = true;
                            rng.Font.Name = "ISOCPEUR"; rng.Font.Size = 10;
                            rng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                            rng.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                        }

                        // I «Примечания» по блоку (пока пусто)
                        var rngNote = ws.Range[ws.Cells[curRow, 9], ws.Cells[blockEnd, 9]];
                        rngNote.Merge();
                        rngNote.Value2 = "";
                        rngNote.WrapText = true; rngNote.ShrinkToFit = true;
                        rngNote.Font.Name = "ISOCPEUR"; rngNote.Font.Size = 10;
                        rngNote.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        rngNote.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                        // строки данных
                        for (int i = 0; i < rows; i++)
                        {
                            int rowIdx = curRow + i;

                            if (i < ceil.Count)
                            {
                                ws.Cells[rowIdx, 3].Value2 = ceil[i].Key;
                                ws.Cells[rowIdx, 4].Value2 = ceil[i].Value;
                            }
                            if (i < walls.Count)
                            {
                                ws.Cells[rowIdx, 5].Value2 = walls[i].Key;
                                ws.Cells[rowIdx, 6].Value2 = walls[i].Value;
                            }
                            if (i < cols.Count)
                            {
                                ws.Cells[rowIdx, 7].Value2 = cols[i].Key;   // уже без префикса
                                ws.Cells[rowIdx, 8].Value2 = cols[i].Value;
                            }

                            StyleDataRow_TriplePairs(ws, rowIdx);
                        }

                        ws.Range[ws.Cells[curRow, 1], ws.Cells[blockEnd, 9]].EntireRow.AutoFit();
                        curRow = blockEnd + 1;
                    }

                    // рамка + автоподгон
                    var border = ws.Range[ws.Cells[2, 1], ws.Cells[Math.Max(curRow - 1, 2), 9]];
                    border.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    border.Borders.Weight = Excel.XlBorderWeight.xlThin;

                    var used = ws.Range[ws.Cells[1, 1], ws.Cells[Math.Max(curRow - 1, 1), 9]];
                    used.WrapText = true;
                    used.Rows.AutoFit();

                    var dlg = new System.Windows.Forms.SaveFileDialog { Filter = "Excel files (*.xlsx)|*.xlsx" };
                    if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                        wb.SaveAs(dlg.FileName);

                    wb.Close(false);
                    xl.Quit();
                    GC.Collect();
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

        /* ------------------- helpers (форматирование) -------------------------------- */
        void FormatHeader8cols_BlockCeiling(Excel.Worksheet ws)
        {
            // ===== Строка 1: общий заголовок =====
            var top = ws.Range[ws.Cells[1, 1], ws.Cells[1, 8]];
            top.Merge();
            top.Value2 = "Ведомость отделки";
            top.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            top.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            top.WrapText = true;
            top.Font.Name = "ISOCPEUR";
            top.Font.Size = 10;
            top.Font.Bold = true;

            // ===== Строка 2 =====
            var rngNum = ws.Range[ws.Cells[2, 1], ws.Cells[3, 1]]; // A
            var rngName = ws.Range[ws.Cells[2, 2], ws.Cells[3, 2]]; // B
            var rngTypes = ws.Range[ws.Cells[2, 3], ws.Cells[2, 7]]; // C–G
            var rngNote = ws.Range[ws.Cells[2, 8], ws.Cells[3, 8]]; // H

            rngNum.Merge(); rngNum.Value2 = "Номера помещений";
            rngName.Merge(); rngName.Value2 = "Наименования помещений";
            rngTypes.Merge(); rngTypes.Value2 = "Типы отделки помещений";
            rngNote.Merge(); rngNote.Value2 = "Примечания";

            // ===== Строка 3: подписи внутри блока C–G =====
            ws.Cells[3, 3].Value2 = "Отделка потолка";
            ws.Cells[3, 4].Value2 = "Отделка стен или перегородок";
            ws.Cells[3, 5].Value2 = "Площ. м2";
            ws.Cells[3, 6].Value2 = "Отделка колонн";
            ws.Cells[3, 7].Value2 = "Площ. м2";

            // Общий стиль для всей шапки
            var rng23 = ws.Range[ws.Cells[2, 1], ws.Cells[3, 8]];
            rng23.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            rng23.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            rng23.WrapText = true;
            rng23.Font.Name = "ISOCPEUR";
            rng23.Font.Size = 10;

            // Левое выравнивание только для текстовых столбцов в подзаголовках
            ws.Cells[3, 3].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;   // потолок
            ws.Cells[3, 4].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;   // стены
            ws.Cells[3, 6].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;   // колонны

            // Центрируем только площади, а примечания остаются как в TriplePairs
            ws.Cells[3, 5].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter; // площ. стен
            ws.Cells[3, 7].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter; // площ. колонн

            // Рамка по всей шапке
            var hdrBorder = ws.Range[ws.Cells[2, 1], ws.Cells[3, 8]];
            hdrBorder.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            hdrBorder.Borders.Weight = Excel.XlBorderWeight.xlThin;
        }
        void StyleDataRow_BlockCeiling(Excel.Worksheet ws, int row)
        {
            // описания: C, D, F, H
            foreach (int col in new[] { 3, 4, 6, 8 })
            {
                ws.Cells[row, col].WrapText = true;
                ws.Cells[row, col].Font.Name = "ISOCPEUR";
                ws.Cells[row, col].Font.Size = 10;
                ws.Cells[row, col].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                ws.Cells[row, col].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            }
            // площади: E, G
            foreach (int col in new[] { 5, 7 })
            {
                ws.Cells[row, col].Font.Name = "ISOCPEUR";
                ws.Cells[row, col].Font.Size = 10;
                ws.Cells[row, col].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                ws.Cells[row, col].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            }
        }
        void FormatHeader9cols_TriplePairs(Excel.Worksheet ws)
        {
            // строка 1
            var top = ws.Range[ws.Cells[1, 1], ws.Cells[1, 9]];
            top.Merge();
            top.Value2 = "Ведомость отделки";
            top.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            top.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            top.WrapText = true;
            top.Font.Name = "ISOCPEUR";
            top.Font.Size = 10;
            top.Font.Bold = true;

            // строка 2 (A,B – вниз; C..H – объединённый блок типов; I – Примечания вниз)
            var rngNum = ws.Range[ws.Cells[2, 1], ws.Cells[3, 1]];
            var rngName = ws.Range[ws.Cells[2, 2], ws.Cells[3, 2]];
            var rngType = ws.Range[ws.Cells[2, 3], ws.Cells[2, 8]];
            var rngNote = ws.Range[ws.Cells[2, 9], ws.Cells[3, 9]];

            rngNum.Merge(); rngNum.Value2 = "Номера помещений";
            rngName.Merge(); rngName.Value2 = "Наименования помещений";
            rngType.Merge(); rngType.Value2 = "Типы отделки помещений";
            rngNote.Merge(); rngNote.Value2 = "Примечания";

            // строка 3 — подписи пар
            ws.Cells[3, 3].Value2 = "Отделка потолка";
            ws.Cells[3, 4].Value2 = "Площ. м2";
            ws.Cells[3, 5].Value2 = "Отделка стен или перегородок";
            ws.Cells[3, 6].Value2 = "Площ. м2";
            ws.Cells[3, 7].Value2 = "Отделка колонн";
            ws.Cells[3, 8].Value2 = "Площ. м2";

            var rng23 = ws.Range[ws.Cells[2, 1], ws.Cells[3, 9]];
            rng23.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            rng23.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            rng23.WrapText = true;
            rng23.Font.Name = "ISOCPEUR";
            rng23.Font.Size = 10;
        }
        void StyleDataRow_TriplePairs(Excel.Worksheet ws, int row)
        {
            // Описания: C, E, G
            foreach (int col in new[] { 3, 5, 7 })
            {
                ws.Cells[row, col].WrapText = true;
                ws.Cells[row, col].Font.Name = "ISOCPEUR";
                ws.Cells[row, col].Font.Size = 10;
                ws.Cells[row, col].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                ws.Cells[row, col].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            }

            // Площади: D, F, H — без установки NumberFormat, чтобы не ловить ошибку
            foreach (int col in new[] { 4, 6, 8 })
            {
                ws.Cells[row, col].Font.Name = "ISOCPEUR";
                ws.Cells[row, col].Font.Size = 10;
                ws.Cells[row, col].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                ws.Cells[row, col].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            }
        }

        public class WallTypeIdComparer : IEqualityComparer<WallType>
        {
            public bool Equals(WallType x, WallType y) { return x.Id == y.Id; }
            public int GetHashCode(WallType obj) { return obj.Id.GetHashCode(); }
        }
        void AccumulateArea(
            Document doc,
            List<WallType> wallTypes,
            string roomNumber,
            Dictionary<string, double> itemData,
            Guid elemData,
            Guid roombookRoomNumber,
            bool isColumn)
        {
            foreach (WallType wallType in wallTypes)
            {
                var wallsOfType = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_Walls)
                    .OfClass(typeof(Wall))
                    .WhereElementIsNotElementType()
                    .Cast<Wall>()
                    .Where(w => w.WallType.Id == wallType.Id
                             && w.get_Parameter(roombookRoomNumber) != null
                             && w.get_Parameter(roombookRoomNumber).AsString() == roomNumber)
                    .ToList();

                double wallArea = 0;
                foreach (Wall wall in wallsOfType)
                {
#if R2019 || R2020 || R2021
                    wallArea += UnitUtils.ConvertFromInternalUnits(
                        wall.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsDouble(),
                        DisplayUnitType.DUT_SQUARE_METERS);
#else
                    wallArea += UnitUtils.ConvertFromInternalUnits(
                        wall.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsDouble(),
                        UnitTypeId.SquareMeters);
#endif
                }

                // Имя типа отделки
                string name = wallType.get_Parameter(elemData)?.AsString();
                if (string.IsNullOrWhiteSpace(name)) name = "(без имени типа)";

                // КЛЮЧЕВОЕ: для колонн добавляем префикс
                string key = isColumn ? $"Колонны — {name}" : name;

                double val = Math.Round(wallArea, 2);
                if (itemData.ContainsKey(key)) itemData[key] += val;
                else itemData[key] = val;
            }
        }
        void AccumulateArea(
            Document doc,
            IEnumerable<WallType> wallTypes,
            string roomNumber,
            Dictionary<string, double> itemData,
            Guid elemData,
            Guid roombookRoomNumber,
            bool isColumn,
            bool addPrefixForColumns = true)
        {
            if (wallTypes == null) return;

            foreach (WallType wallType in wallTypes)
            {
                // страховка: проверяем, колонный это тип или нет
                string model = wallType.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL)?.AsString() ?? string.Empty;
                bool typeIsColumn = model.IndexOf("Колонны", StringComparison.OrdinalIgnoreCase) >= 0;

                if (isColumn && !typeIsColumn) continue;        // ждём колонны — получили не колонны
                if (!isColumn && typeIsColumn) continue;        // ждём стены — получили колонны

                // стены/колонны данного типа в помещении
                var wallsOfType = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_Walls)
                    .OfClass(typeof(Wall))
                    .WhereElementIsNotElementType()
                    .Cast<Wall>()
                    .Where(w => w.WallType.Id == wallType.Id
                             && w.get_Parameter(roombookRoomNumber) != null
                             && w.get_Parameter(roombookRoomNumber).AsString() == roomNumber)
                    .ToList();

                double wallArea = 0;
                foreach (Wall wall in wallsOfType)
                {
#if R2019 || R2020 || R2021
                    wallArea += UnitUtils.ConvertFromInternalUnits(
                        wall.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsDouble(),
                        DisplayUnitType.DUT_SQUARE_METERS);
#else
                    wallArea += UnitUtils.ConvertFromInternalUnits(
                        wall.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsDouble(),
                        UnitTypeId.SquareMeters);
#endif
                }

                // имя типа отделки
                string name = wallType.get_Parameter(elemData)?.AsString();
                if (string.IsNullOrWhiteSpace(name)) name = "(без имени типа)";

                // префикс оставляем только там, где нужен для совместного словаря
                string key = (isColumn && addPrefixForColumns) ? $"Колонны — {name}" : name;

                double val = Math.Round(wallArea, 2);
                if (itemData.ContainsKey(key)) itemData[key] += val;
                else itemData[key] = val;
            }
        }
        void AccumulateCeilingArea(
            Document doc,
            IEnumerable<CeilingType> types, 
            string roomNumber,
            Dictionary<string, double> target,  
            Guid elemDataParam,              
            Guid roomNumberParam)
        {
            if (target == null) return;

            // Если типов не дали — соберём фактические типы потолков в помещении
            IEnumerable<CeilingType> typeList = types;
            if (typeList == null)
            {
                typeList = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_Ceilings)
                    .OfClass(typeof(Ceiling))
                    .WhereElementIsNotElementType()
                    .Cast<Ceiling>()
                    .Where(c => c.get_Parameter(roomNumberParam) != null &&
                                c.get_Parameter(roomNumberParam).AsString() == roomNumber)
                    .Select(c => doc.GetElement(c.GetTypeId()) as CeilingType)
                    .Where(ct => ct != null &&
                                 ct.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL) != null &&
                                (ct.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString() == "Потолок" ||
                                 ct.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString() == "Потолки"))
                    .Distinct(new ElementTypeIdComparer<CeilingType>())
                    .ToList();
            }

            foreach (var ct in typeList)
            {
                if (ct == null) continue;

                // Все потолки данного типа в этом помещении
                var elems = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_Ceilings)
                    .OfClass(typeof(Ceiling))
                    .WhereElementIsNotElementType()
                    .Cast<Ceiling>()
                    .Where(c => c.GetTypeId() == ct.Id &&
                                c.get_Parameter(roomNumberParam) != null &&
                                c.get_Parameter(roomNumberParam).AsString() == roomNumber)
                    .ToList();

                if (elems.Count == 0) continue;

                double sum = 0.0;
                foreach (var c in elems)
                {
#if R2019 || R2020 || R2021
                    sum += UnitUtils.ConvertFromInternalUnits(
                        c.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsDouble(),
                        DisplayUnitType.DUT_SQUARE_METERS);
#else
                    sum += UnitUtils.ConvertFromInternalUnits(
                        c.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsDouble(),
                        UnitTypeId.SquareMeters);
#endif
                }

                // Ключ — описание типа без префиксов
                string key = ct.get_Parameter(elemDataParam)?.AsString();
                if (string.IsNullOrWhiteSpace(key)) key = ct.Name ?? "(без имени типа)";

                double val = Math.Round(sum, 2);
                if (target.ContainsKey(key)) target[key] += val;
                else target.Add(key, val);
            }
        }
        private void ThreadStartingPoint()
        {
            roomBookToExcelProgressBarWPF = new RoomBookToExcelProgressBarWPF();
            roomBookToExcelProgressBarWPF.Show();
            System.Windows.Threading.Dispatcher.Run();
        }
        private static async Task GetPluginStartInfo()
        {
            // Получаем сборку, в которой выполняется текущий код
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
                // Создание экземпляра класса
                object instance = Activator.CreateInstance(type);

                // Получение метода CollectPluginUsageAsync
                var method = type.GetMethod("CollectPluginUsageAsync");

                if (method != null)
                {
                    // Вызов асинхронного метода через reflection
                    Task task = (Task)method.Invoke(instance, new object[] { assemblyName, assemblyNameRus });
                    await task;  // Ожидание завершения асинхронного метода
                }
            }
        }
    }
}

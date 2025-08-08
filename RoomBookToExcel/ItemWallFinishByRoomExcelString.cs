using System.Collections.Generic;

namespace RoomBookToExcel
{
    public class ItemWallFinishByRoomExcelString
    {
        public string RoomNumber { get; set; }
        public string RoomName { get; set; }
        public Dictionary<string, double> ItemData { get; } = new Dictionary<string, double>();
    }
}

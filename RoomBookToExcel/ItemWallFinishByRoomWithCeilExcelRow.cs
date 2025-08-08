using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RoomBookToExcel
{
    public class ItemWallFinishByRoomWithCeilExcelRow
    {
        public string RoomNumber { get; set; }
        public string RoomName { get; set; }

        // тип -> площадь
        public Dictionary<string, double> CeilingData { get; } = new Dictionary<string, double>();
        public Dictionary<string, double> WallData { get; } = new Dictionary<string, double>();
        public Dictionary<string, double> ColumnData { get; } = new Dictionary<string, double>();
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RoomBookToExcel
{
    class ItemWallFinishByRoomWithCeilComparer : IEqualityComparer<ItemWallFinishByRoomWithCeil>
    {
        public bool Equals(ItemWallFinishByRoomWithCeil x, ItemWallFinishByRoomWithCeil y) => x.Equals(y);
        public int GetHashCode(ItemWallFinishByRoomWithCeil obj) => obj.GetHashCode();
    }
}

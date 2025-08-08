using Autodesk.Revit.DB;
using System.Collections.Generic;
using System.Linq;

namespace RoomBookToExcel
{
    class ItemWallFinishByRoom
    {
        public string RoomNumber { get; set; }
        public string RoomName { get; set; }

        public List<WallType> WallFinishes { get; set; } = new List<WallType>();
        public List<WallType> ColumnFinishes { get; set; } = new List<WallType>();

        /* ---------- Сравнение двух сочетаний ---------- */

        public override bool Equals(object obj)
        {
            var other = obj as ItemWallFinishByRoom;
            if (other == null)
                return false;

            return ListsEqual(WallFinishes, other.WallFinishes) &&
                   ListsEqual(ColumnFinishes, other.ColumnFinishes);
        }

        public override int GetHashCode()
        {
            int hash = 17;
            hash = hash * 31 + GetListHash(WallFinishes);
            hash = hash * 31 + GetListHash(ColumnFinishes);
            return hash;
        }

        /* ---------- Вспомогалки ---------- */

        private static bool ListsEqual(IReadOnlyList<WallType> a, IReadOnlyList<WallType> b)
        {
            if (a.Count != b.Count) return false;
            for (int i = 0; i < a.Count; i++)
                if (a[i].Id != b[i].Id) return false;
            return true;
        }

        private static int GetListHash(IEnumerable<WallType> list) =>
            list.Aggregate(19, (h, t) => h * 31 + t.Id.GetHashCode());
    }
}

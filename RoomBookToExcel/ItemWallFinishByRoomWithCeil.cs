using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RoomBookToExcel
{
    using Autodesk.Revit.DB;
    using System.Collections.Generic;
    using System.Linq;

    class ItemWallFinishByRoomWithCeil
    {
        public string RoomNumber { get; set; }
        public string RoomName { get; set; }

        public List<WallType> WallFinishes { get; } = new List<WallType>();
        public List<WallType> ColumnFinishes { get; } = new List<WallType>();
        public List<CeilingType> CeilingFinishes { get; } = new List<CeilingType>(); // НОВОЕ

        /* ---------- Сравнение двух сочетаний ---------- */
        public override bool Equals(object obj)
        {
            var other = obj as ItemWallFinishByRoomWithCeil;
            if (other == null) return false;

            return ListsEqual(WallFinishes, other.WallFinishes) &&
                   ListsEqual(ColumnFinishes, other.ColumnFinishes) &&
                   ListsEqual(CeilingFinishes, other.CeilingFinishes);
        }

        public override int GetHashCode()
        {
            int hash = 17;
            hash = hash * 31 + GetListHash(WallFinishes);
            hash = hash * 31 + GetListHash(ColumnFinishes);
            hash = hash * 31 + GetListHash(CeilingFinishes);
            return hash;
        }

        /* ---------- Вспомогалки (как у тебя, но обобщим под ElementType) ---------- */
        private static bool ListsEqual<T>(IReadOnlyList<T> a, IReadOnlyList<T> b) where T : ElementType
        {
            if (a.Count != b.Count) return false;
            for (int i = 0; i < a.Count; i++)
                if (a[i].Id != b[i].Id) return false;
            return true;
        }

        private static int GetListHash<T>(IEnumerable<T> list) where T : ElementType =>
            list.Aggregate(19, (h, t) => h * 31 + t.Id.GetHashCode());
    }
}

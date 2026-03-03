using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RoomBookToExcel
{
    class ElementTypeIdComparer<T> : IEqualityComparer<T> where T : ElementType
    {
        public bool Equals(T x, T y)
        {
            if (ReferenceEquals(x, y)) return true;
            if (x is null || y is null) return false;
            return x.Id.Equals(y.Id);
        }

        public int GetHashCode(T obj)
        {
            if (obj is null) return 0;
            return obj.Id.GetHashCode();
        }
    }
}

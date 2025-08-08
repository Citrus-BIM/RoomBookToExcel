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
        public bool Equals(T x, T y) => x?.Id.IntegerValue == y?.Id.IntegerValue;
        public int GetHashCode(T obj) => obj?.Id.IntegerValue.GetHashCode() ?? 0;
    }
}

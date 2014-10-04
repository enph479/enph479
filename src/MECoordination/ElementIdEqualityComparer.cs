using Autodesk.Revit.DB;
using System.Collections.Generic;

namespace ElectricalToolSuite.MECoordination
{
    public class ElementIdEqualityComparer<T> : IEqualityComparer<T> where T : Element
    {
        public bool Equals(T x, T y)
        {
            return x.Id == y.Id;
        }

        public int GetHashCode(T obj)
        {
            return obj.GetHashCode();
        }
    }
}
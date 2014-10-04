using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ElectricalToolSuite.MECoordination
{
    public class ElementNameEqualityComparer<T> : IEqualityComparer<T> where T : Element
    {
        public bool Equals(T x, T y)
        {
            return x.Name == y.Name;
        }

        public int GetHashCode(T obj)
        {
            return obj.GetHashCode();
        }
    }
}
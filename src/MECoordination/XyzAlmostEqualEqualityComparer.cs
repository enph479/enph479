using Autodesk.Revit.DB;
using System.Collections.Generic;

namespace ElectricalToolSuite.MECoordination
{
    class XyzAlmostEqualEqualityComparer : IEqualityComparer<XYZ>
    {
        public bool Equals(XYZ x, XYZ y)
        {
            return x.IsAlmostEqualTo(y);
        }

        public int GetHashCode(XYZ obj)
        {
            return obj.GetHashCode();
        }
    }
}

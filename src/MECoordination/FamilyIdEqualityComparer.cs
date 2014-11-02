using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace ElectricalToolSuite.MECoordination
{
    public class FamilyIdEqualityComparer : IEqualityComparer<Family>
    {
        public bool Equals(Family x, Family y)
        {
            return x.Id.IntegerValue == y.Id.IntegerValue;
        }

        public int GetHashCode(Family obj)
        {
            return obj.Id.GetHashCode();
        }
    }
}
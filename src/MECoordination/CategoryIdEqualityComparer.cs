using Autodesk.Revit.DB;
using System.Collections.Generic;

namespace ElectricalToolSuite.MECoordination
{
    internal class CategoryIdEqualityComparer : IEqualityComparer<Category>
    {
        public bool Equals(Category x, Category y)
        {
            return x.Id.IntegerValue == y.Id.IntegerValue;
        }

        public int GetHashCode(Category obj)
        {
            return obj.Id.GetHashCode();
        }
    }
}

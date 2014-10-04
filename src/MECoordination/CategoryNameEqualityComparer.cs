using Autodesk.Revit.DB;
using System.Collections.Generic;

namespace ElectricalToolSuite.MECoordination
{
    public class CategoryNameEqualityComparer : IEqualityComparer<Category>
    {
        public bool Equals(Category x, Category y)
        {
            return x.Name == y.Name;
        }

        public int GetHashCode(Category obj)
        {
            return obj.GetHashCode();
        }
    }
}
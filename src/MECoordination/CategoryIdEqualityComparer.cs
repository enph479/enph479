using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;

namespace ElectricalToolSuite.MECoordination
{
    internal class CategoryIdEqualityComparer : IEqualityComparer<Category>
    {
        public bool Equals(Category x, Category y)
        {
            return x.Id.Equals(y.Id);
        }

        public int GetHashCode(Category obj)
        {
            return obj.Id.GetHashCode();
        }
    }
}

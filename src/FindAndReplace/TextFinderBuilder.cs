using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ElectricalToolSuite.FindAndReplace
{
    class TextFinderBuilder
    {
        public static TextFinder BuildTextFinder(FinderSettings finderSettings)
        {
            return new TextFinder(finderSettings._searchText);
        }

    }
}

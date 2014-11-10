using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace Redbolts.UI.Common.Composition
{
    public interface IRibbonItemMetaData
    {
        int PanelIndex { get; }
        string TabName { get; }
        string PanelName { get; }

        string LongDescription { get; }
        string Name { get; }
        string Tooltip { get; }
        string TooltipImage { get; }

        bool Visible { get; }
        bool Enabled { get; }
    }
}

using System.Drawing;

namespace ElectricalToolSuite.ScheduleImport
{
    public static class ColorExtensions
    {
        public static double ToDouble(this Color color)
        {
            uint rgb = color.B;
            rgb = rgb << 8;
            rgb += color.G;
            rgb = rgb << 8;
            rgb += color.R;
            return rgb;
        }
    }
}

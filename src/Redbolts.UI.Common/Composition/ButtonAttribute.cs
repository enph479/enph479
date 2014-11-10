using System;
using System.ComponentModel.Composition;
using Autodesk.Revit.UI;

namespace Redbolts.UI.Common.Composition
{
    [MetadataAttribute]
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false)]
    public class ButtonAttribute : ExportAttribute, IButtonMetaData
    {
        public int PanelIndex { get; set; }

        public string TabName { get; set; }

        public string PanelName { get; set; }

        public string LongDescription { get; set; }

        public string Name { get; set; }

        public string Tooltip { get; set; }

        public string TooltipImage { get; set; }

        public bool Visible { get; set; }

        public bool Enabled { get; set; }

        public string Text { get; set; }

        public string Image { get; set; }

        public string LargeImage { get; set; }

        public ButtonAttribute(string tabName, string panelName, string name, string buttonText)
            : base(typeof(IExternalCommand))
        {
            PanelIndex = 0;
            TabName = tabName;
            PanelName = panelName;
            Name = name;
            Text = buttonText;
            Enabled = true;
            Visible = true;
        }
    }
}
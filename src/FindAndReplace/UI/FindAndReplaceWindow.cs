using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Windows.Documents;
using Autodesk.Revit.DB;
using ElectricalToolSuite.FindAndReplace.UI;
using Form = System.Windows.Forms.Form;

namespace ElectricalToolSuite.FindAndReplace
{
    public partial class FindAndReplaceWindow : Form
    {
        public bool NotCancelled { set; get; }
        private readonly FinderSettings _finderSettings;
        private List<ViewSelectorDto> _searchableViews;
        public FindAndReplaceWindow(FilteredElementCollector allViews, View activeView)
        {
            InitializeComponent();
            _finderSettings = new FinderSettings();
            _searchableViews = new List<ViewSelectorDto>();

            foreach (View view in allViews.OfType<ViewPlan>())
            {
                _searchableViews.Add(view.Id.Equals(activeView.Id)
                    ? new ViewSelectorDto(true, view)
                    : new ViewSelectorDto(false, view));
            }
        }

        private void FindTextBox_TextChanged(object sender, EventArgs e)
        {
            _finderSettings.SearchText = FindTextBox.Text;
        }

        private void CaseSensitiveCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            _finderSettings.CaseSensitive = CaseSensitiveCheckBox.Checked;
        }

        private void SelectedViewsRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            if (SelectedViewsRadioButton.Checked)
            {
                _finderSettings.SearchViewFilter = SearchViewSettings.SelectedView;
                var temp = new SelectedViewResults(_searchableViews);
                temp.Show();
            }
        }

        private void WholeWordsCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            _finderSettings.WholeWords = WholeWordsCheckBox.Checked;
        }

        private void HiddenElementCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            _finderSettings.HiddenElements = HiddenElementCheckBox.Checked;
        }

        private void FindButton_Click(object sender, EventArgs e)
        {
            NotCancelled = true;
            Close();
        }

        private void CancelButton_Click(object sender, EventArgs e)
        {
            NotCancelled = false;
            Close();
        }

        private void CurrentViewRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            if (CurrentViewRadioButton.Checked)
            {
                _finderSettings.SearchViewFilter = SearchViewSettings.CurrentView;                
            }
        }

        private void EntireProjectRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            if (EntireProjectRadioButton.Checked)
            {
                _finderSettings.SearchViewFilter = SearchViewSettings.AllViews;
                foreach (ViewSelectorDto view in _searchableViews)
                {
                    view.IsChecked = true;
                }
            }
        }

        public FinderSettings GetFinderSettings()
        {
            return _finderSettings;
        }

        public List<ViewSelectorDto> GetSearchableViews()
        {
            var removeList = new List<ViewSelectorDto>(_searchableViews);
            foreach (ViewSelectorDto view in removeList)
            {
                if (!view.IsChecked)
                {
                    _searchableViews.Remove(view);
                }
            }
            return _searchableViews;
        }

        private void FindLabel_Click(object sender, EventArgs e)
        {

        }

        private void FindAndReplaceWindow_Load(object sender, EventArgs e)
        {

        }
    }
}

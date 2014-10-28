using System;
using System.Diagnostics;
using System.Windows.Forms;

namespace ElectricalToolSuite.FindAndReplace
{
    public partial class FindAndReplaceWindow : Form
    {

        private readonly FinderSettings _finderSettings; 
        public FindAndReplaceWindow()
        {
            InitializeComponent();
            _finderSettings = new FinderSettings();
        }

        private void FindTextBox_TextChanged(object sender, EventArgs e)
        {
            _finderSettings.SearchText = FindTextBox.Text;
        }

        private void ReplaceTextBox_TextChanged(object sender, EventArgs e)
        {
            _finderSettings.ReplaceText = ReplaceTextBox.Text;
        }

        private void CaseSensitiveCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            _finderSettings.CaseSensitive = CaseSensitiveCheckBox.Checked;
        }

        private void SelectedViewsRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            _finderSettings.SearchViewFilter = SearchViewSettings.SelectedView;
            //TODO: for selected views make it so a new dialog box appears to selected the selected views 
        }

        private void WholeWordsCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            _finderSettings.WholeWords = WholeWordsCheckBox.Checked;
        }

        private void HiddenElementCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            _finderSettings.HiddenElements = HiddenElementCheckBox.Checked;
        }

        private void DimensionCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            _finderSettings.DimensionText = DimensionCheckBox.Checked;
        }

        private void TableTextCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            _finderSettings.IncludeTableText = TableTextCheckBox.Checked;
        }

        private void FamilyPropertiesCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            _finderSettings.FamilyProperties = FamilyPropertiesCheckBox.Checked;
        }

        private void FindButton_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void ReplaceButton_Click(object sender, EventArgs e)
        {
            //TODO: Actually implement replace functionality
            Close();
        }

        private void CancelButton_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void CurrentViewRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            _finderSettings.SearchViewFilter = SearchViewSettings.CurrentView;
        }

        private void EntireProjectRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            _finderSettings.SearchViewFilter = SearchViewSettings.AllViews;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        public FinderSettings GetFinderSettings()
        {
            return _finderSettings;
        }
        private void tableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {
        }

    }
}

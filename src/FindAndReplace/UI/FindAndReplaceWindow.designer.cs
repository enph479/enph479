using System.Windows.Forms;

namespace ElectricalToolSuite.FindAndReplace
{
    partial class FindAndReplaceWindow
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.TableLayoutPanel = new System.Windows.Forms.TableLayoutPanel();
            this.CaseSensitiveCheckBox = new System.Windows.Forms.CheckBox();
            this.WholeWordsCheckBox = new System.Windows.Forms.CheckBox();
            this.ScopeLabel = new System.Windows.Forms.Label();
            this.HiddenElementCheckBox = new System.Windows.Forms.CheckBox();
            this.FindButton = new System.Windows.Forms.Button();
            this.FindTextBox = new System.Windows.Forms.TextBox();
            this.FindLabel = new System.Windows.Forms.Label();
            this.CurrentViewRadioButton = new System.Windows.Forms.RadioButton();
            this.EntireProjectRadioButton = new System.Windows.Forms.RadioButton();
            this.SelectedViewsRadioButton = new System.Windows.Forms.RadioButton();
            this.CancelButton = new System.Windows.Forms.Button();
            this.TableLayoutPanel.SuspendLayout();
            this.SuspendLayout();
            // 
            // TableLayoutPanel
            // 
            this.TableLayoutPanel.AutoSize = true;
            this.TableLayoutPanel.ColumnCount = 3;
            this.TableLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.TableLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 246F));
            this.TableLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 9F));
            this.TableLayoutPanel.Controls.Add(this.WholeWordsCheckBox, 0, 3);
            this.TableLayoutPanel.Controls.Add(this.FindTextBox, 1, 0);
            this.TableLayoutPanel.Controls.Add(this.FindLabel, 0, 0);
            this.TableLayoutPanel.Controls.Add(this.CaseSensitiveCheckBox, 0, 2);
            this.TableLayoutPanel.Controls.Add(this.CancelButton, 1, 11);
            this.TableLayoutPanel.Controls.Add(this.FindButton, 0, 11);
            this.TableLayoutPanel.Controls.Add(this.EntireProjectRadioButton, 0, 9);
            this.TableLayoutPanel.Controls.Add(this.SelectedViewsRadioButton, 0, 8);
            this.TableLayoutPanel.Controls.Add(this.CurrentViewRadioButton, 0, 7);
            this.TableLayoutPanel.Controls.Add(this.HiddenElementCheckBox, 1, 7);
            this.TableLayoutPanel.Controls.Add(this.ScopeLabel, 0, 5);
            this.TableLayoutPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.TableLayoutPanel.Location = new System.Drawing.Point(0, 0);
            this.TableLayoutPanel.Name = "TableLayoutPanel";
            this.TableLayoutPanel.RowCount = 12;
            this.TableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.TableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 22F));
            this.TableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.TableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 24F));
            this.TableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 22F));
            this.TableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 25F));
            this.TableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.TableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 24F));
            this.TableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 26F));
            this.TableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 28F));
            this.TableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 31F));
            this.TableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 14F));
            this.TableLayoutPanel.Size = new System.Drawing.Size(394, 285);
            this.TableLayoutPanel.TabIndex = 0;
            // 
            // CaseSensitiveCheckBox
            // 
            this.CaseSensitiveCheckBox.AutoSize = true;
            this.CaseSensitiveCheckBox.Location = new System.Drawing.Point(3, 51);
            this.CaseSensitiveCheckBox.Name = "CaseSensitiveCheckBox";
            this.CaseSensitiveCheckBox.Size = new System.Drawing.Size(94, 17);
            this.CaseSensitiveCheckBox.TabIndex = 2;
            this.CaseSensitiveCheckBox.Text = "Case sensitive";
            this.CaseSensitiveCheckBox.UseVisualStyleBackColor = true;
            this.CaseSensitiveCheckBox.CheckedChanged += new System.EventHandler(this.CaseSensitiveCheckBox_CheckedChanged);
            // 
            // WholeWordsCheckBox
            // 
            this.WholeWordsCheckBox.AutoSize = true;
            this.WholeWordsCheckBox.Location = new System.Drawing.Point(3, 74);
            this.WholeWordsCheckBox.Name = "WholeWordsCheckBox";
            this.WholeWordsCheckBox.Size = new System.Drawing.Size(110, 17);
            this.WholeWordsCheckBox.TabIndex = 3;
            this.WholeWordsCheckBox.Text = "Whole words only";
            this.WholeWordsCheckBox.UseVisualStyleBackColor = true;
            this.WholeWordsCheckBox.CheckedChanged += new System.EventHandler(this.WholeWordsCheckBox_CheckedChanged);
            // 
            // ScopeLabel
            // 
            this.ScopeLabel.AutoSize = true;
            this.ScopeLabel.Location = new System.Drawing.Point(3, 117);
            this.ScopeLabel.Name = "ScopeLabel";
            this.ScopeLabel.Size = new System.Drawing.Size(38, 13);
            this.ScopeLabel.TabIndex = 8;
            this.ScopeLabel.Text = "Scope";
            // 
            // HiddenElementCheckBox
            // 
            this.HiddenElementCheckBox.AutoSize = true;
            this.HiddenElementCheckBox.Location = new System.Drawing.Point(142, 145);
            this.HiddenElementCheckBox.Name = "HiddenElementCheckBox";
            this.HiddenElementCheckBox.Size = new System.Drawing.Size(141, 17);
            this.HiddenElementCheckBox.TabIndex = 9;
            this.HiddenElementCheckBox.Text = "Include hidden elements";
            this.HiddenElementCheckBox.UseVisualStyleBackColor = true;
            this.HiddenElementCheckBox.CheckedChanged += new System.EventHandler(this.HiddenElementCheckBox_CheckedChanged);
            // 
            // FindButton
            // 
            this.FindButton.Location = new System.Drawing.Point(3, 254);
            this.FindButton.Name = "FindButton";
            this.FindButton.Size = new System.Drawing.Size(75, 23);
            this.FindButton.TabIndex = 13;
            this.FindButton.Text = "Find";
            this.FindButton.UseVisualStyleBackColor = true;
            this.FindButton.Click += new System.EventHandler(this.FindButton_Click);
            // 
            // FindTextBox
            // 
            this.FindTextBox.Location = new System.Drawing.Point(142, 3);
            this.FindTextBox.Name = "FindTextBox";
            this.FindTextBox.Size = new System.Drawing.Size(223, 20);
            this.FindTextBox.TabIndex = 0;
            this.FindTextBox.TextChanged += new System.EventHandler(this.FindTextBox_TextChanged);
            // 
            // FindLabel
            // 
            this.FindLabel.AutoSize = true;
            this.FindLabel.Location = new System.Drawing.Point(3, 0);
            this.FindLabel.Name = "FindLabel";
            this.FindLabel.Size = new System.Drawing.Size(47, 13);
            this.FindLabel.TabIndex = 16;
            this.FindLabel.Text = "Find text";
            this.FindLabel.Click += new System.EventHandler(this.FindLabel_Click);
            // 
            // CurrentViewRadioButton
            // 
            this.CurrentViewRadioButton.AutoSize = true;
            this.CurrentViewRadioButton.Checked = true;
            this.CurrentViewRadioButton.Location = new System.Drawing.Point(3, 145);
            this.CurrentViewRadioButton.Name = "CurrentViewRadioButton";
            this.CurrentViewRadioButton.Size = new System.Drawing.Size(115, 17);
            this.CurrentViewRadioButton.TabIndex = 4;
            this.CurrentViewRadioButton.TabStop = true;
            this.CurrentViewRadioButton.Text = "Current view/sheet";
            this.CurrentViewRadioButton.UseVisualStyleBackColor = true;
            this.CurrentViewRadioButton.CheckedChanged += new System.EventHandler(this.CurrentViewRadioButton_CheckedChanged);
            // 
            // EntireProjectRadioButton
            // 
            this.EntireProjectRadioButton.AutoSize = true;
            this.EntireProjectRadioButton.Location = new System.Drawing.Point(3, 195);
            this.EntireProjectRadioButton.Name = "EntireProjectRadioButton";
            this.EntireProjectRadioButton.Size = new System.Drawing.Size(87, 17);
            this.EntireProjectRadioButton.TabIndex = 7;
            this.EntireProjectRadioButton.Text = "Entire project";
            this.EntireProjectRadioButton.UseVisualStyleBackColor = true;
            this.EntireProjectRadioButton.CheckedChanged += new System.EventHandler(this.EntireProjectRadioButton_CheckedChanged);
            // 
            // SelectedViewsRadioButton
            // 
            this.SelectedViewsRadioButton.AutoSize = true;
            this.SelectedViewsRadioButton.Location = new System.Drawing.Point(3, 169);
            this.SelectedViewsRadioButton.Name = "SelectedViewsRadioButton";
            this.SelectedViewsRadioButton.Size = new System.Drawing.Size(133, 17);
            this.SelectedViewsRadioButton.TabIndex = 5;
            this.SelectedViewsRadioButton.Text = "Selected views/sheets";
            this.SelectedViewsRadioButton.UseVisualStyleBackColor = true;
            this.SelectedViewsRadioButton.CheckedChanged += new System.EventHandler(this.SelectedViewsRadioButton_CheckedChanged);
            // 
            // CancelButton
            // 
            this.CancelButton.Location = new System.Drawing.Point(142, 254);
            this.CancelButton.Name = "CancelButton";
            this.CancelButton.Size = new System.Drawing.Size(75, 23);
            this.CancelButton.TabIndex = 15;
            this.CancelButton.Text = "Cancel";
            this.CancelButton.UseVisualStyleBackColor = true;
            this.CancelButton.Click += new System.EventHandler(this.CancelButton_Click);
            // 
            // FindAndReplaceWindow
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(394, 285);
            this.Controls.Add(this.TableLayoutPanel);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FindAndReplaceWindow";
            this.Text = "Find Tool";
            this.Load += new System.EventHandler(this.FindAndReplaceWindow_Load);
            this.TableLayoutPanel.ResumeLayout(false);
            this.TableLayoutPanel.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel TableLayoutPanel;
        private System.Windows.Forms.TextBox FindTextBox;
        private System.Windows.Forms.CheckBox CaseSensitiveCheckBox;
        private System.Windows.Forms.CheckBox WholeWordsCheckBox;
        private System.Windows.Forms.RadioButton CurrentViewRadioButton;
        private System.Windows.Forms.RadioButton SelectedViewsRadioButton;
        private System.Windows.Forms.RadioButton EntireProjectRadioButton;
        private System.Windows.Forms.Label ScopeLabel;
        private System.Windows.Forms.CheckBox HiddenElementCheckBox;
        private System.Windows.Forms.Button FindButton;
        private System.Windows.Forms.Button CancelButton;
        private System.Windows.Forms.Label FindLabel;

    }
}


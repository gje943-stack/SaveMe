
namespace src
{
    partial class MainFormView
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tlpMainLayout = new System.Windows.Forms.TableLayoutPanel();
            this.panelButtonsContainer = new System.Windows.Forms.Panel();
            this.dropdownAutoSaveFrequencies = new System.Windows.Forms.ListBox();
            this.labelSaveFrequency = new System.Windows.Forms.Label();
            this.btnSaveAll = new System.Windows.Forms.Button();
            this.btnSaveSelected = new System.Windows.Forms.Button();
            this.btnSelectNone = new System.Windows.Forms.Button();
            this.btnSelectAll = new System.Windows.Forms.Button();
            this.listAllApplications = new System.Windows.Forms.CheckedListBox();
            this.tlpMainLayout.SuspendLayout();
            this.panelButtonsContainer.SuspendLayout();
            this.SuspendLayout();
            // 
            // tlpMainLayout
            // 
            this.tlpMainLayout.ColumnCount = 1;
            this.tlpMainLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tlpMainLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tlpMainLayout.Controls.Add(this.panelButtonsContainer, 0, 0);
            this.tlpMainLayout.Controls.Add(this.listAllApplications, 0, 1);
            this.tlpMainLayout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tlpMainLayout.Location = new System.Drawing.Point(0, 0);
            this.tlpMainLayout.Name = "tlpMainLayout";
            this.tlpMainLayout.RowCount = 2;
            this.tlpMainLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 20.64897F));
            this.tlpMainLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 79.35104F));
            this.tlpMainLayout.Size = new System.Drawing.Size(368, 339);
            this.tlpMainLayout.TabIndex = 0;
            // 
            // panelButtonsContainer
            // 
            this.panelButtonsContainer.Controls.Add(this.dropdownAutoSaveFrequencies);
            this.panelButtonsContainer.Controls.Add(this.labelSaveFrequency);
            this.panelButtonsContainer.Controls.Add(this.btnSaveAll);
            this.panelButtonsContainer.Controls.Add(this.btnSaveSelected);
            this.panelButtonsContainer.Controls.Add(this.btnSelectNone);
            this.panelButtonsContainer.Controls.Add(this.btnSelectAll);
            this.panelButtonsContainer.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelButtonsContainer.Location = new System.Drawing.Point(3, 3);
            this.panelButtonsContainer.Name = "panelButtonsContainer";
            this.panelButtonsContainer.Size = new System.Drawing.Size(362, 63);
            this.panelButtonsContainer.TabIndex = 1;
            // 
            // dropdownAutoSaveFrequencies
            // 
            this.dropdownAutoSaveFrequencies.FormattingEnabled = true;
            this.dropdownAutoSaveFrequencies.ItemHeight = 15;
            this.dropdownAutoSaveFrequencies.Items.AddRange(new object[] {
            "10 Minutes",
            "30 Minute",
            "1 Hour"});
            this.dropdownAutoSaveFrequencies.Location = new System.Drawing.Point(200, 37);
            this.dropdownAutoSaveFrequencies.Name = "dropdownAutoSaveFrequencies";
            this.dropdownAutoSaveFrequencies.Size = new System.Drawing.Size(76, 19);
            this.dropdownAutoSaveFrequencies.TabIndex = 6;
            // 
            // labelSaveFrequency
            // 
            this.labelSaveFrequency.AutoSize = true;
            this.labelSaveFrequency.Location = new System.Drawing.Point(90, 37);
            this.labelSaveFrequency.Name = "labelSaveFrequency";
            this.labelSaveFrequency.Size = new System.Drawing.Size(95, 15);
            this.labelSaveFrequency.TabIndex = 5;
            this.labelSaveFrequency.Text = "Auto-save every:";
            // 
            // btnSaveAll
            // 
            this.btnSaveAll.Location = new System.Drawing.Point(282, 0);
            this.btnSaveAll.Name = "btnSaveAll";
            this.btnSaveAll.Size = new System.Drawing.Size(80, 31);
            this.btnSaveAll.TabIndex = 3;
            this.btnSaveAll.Text = "Save All";
            this.btnSaveAll.UseVisualStyleBackColor = true;
            // 
            // btnSaveSelected
            // 
            this.btnSaveSelected.Location = new System.Drawing.Point(186, 0);
            this.btnSaveSelected.Name = "btnSaveSelected";
            this.btnSaveSelected.Size = new System.Drawing.Size(90, 31);
            this.btnSaveSelected.TabIndex = 2;
            this.btnSaveSelected.Text = "Save Selected";
            this.btnSaveSelected.UseVisualStyleBackColor = true;
            // 
            // btnSelectNone
            // 
            this.btnSelectNone.Location = new System.Drawing.Point(90, 0);
            this.btnSelectNone.Name = "btnSelectNone";
            this.btnSelectNone.Size = new System.Drawing.Size(90, 31);
            this.btnSelectNone.TabIndex = 1;
            this.btnSelectNone.Text = "Select None";
            this.btnSelectNone.UseVisualStyleBackColor = true;
            // 
            // btnSelectAll
            // 
            this.btnSelectAll.Location = new System.Drawing.Point(0, 0);
            this.btnSelectAll.Name = "btnSelectAll";
            this.btnSelectAll.Size = new System.Drawing.Size(84, 31);
            this.btnSelectAll.TabIndex = 0;
            this.btnSelectAll.Text = "Select All";
            this.btnSelectAll.UseVisualStyleBackColor = true;
            // 
            // listAllApplications
            // 
            this.listAllApplications.Dock = System.Windows.Forms.DockStyle.Fill;
            this.listAllApplications.FormattingEnabled = true;
            this.listAllApplications.Location = new System.Drawing.Point(3, 72);
            this.listAllApplications.Name = "listAllApplications";
            this.listAllApplications.Size = new System.Drawing.Size(362, 264);
            this.listAllApplications.TabIndex = 0;
            // 
            // MainFormView
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(368, 339);
            this.Controls.Add(this.tlpMainLayout);
            this.Name = "MainFormView";
            this.Text = "Form1";
            this.tlpMainLayout.ResumeLayout(false);
            this.panelButtonsContainer.ResumeLayout(false);
            this.panelButtonsContainer.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tlpMainLayout;
        private System.Windows.Forms.CheckedListBox listAllApplications;
        private System.Windows.Forms.Panel panelButtonsContainer;
        private System.Windows.Forms.Button btnSaveAll;
        private System.Windows.Forms.Button btnSaveSelected;
        private System.Windows.Forms.Button btnSelectNone;
        private System.Windows.Forms.Button btnSelectAll;
        private System.Windows.Forms.Label labelSaveFrequency;
        internal System.Windows.Forms.ListBox dropdownAutoSaveFrequencies;
    }
}


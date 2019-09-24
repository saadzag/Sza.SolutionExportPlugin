using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using McTools.Xrm.Connection;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Sza.SolutionExportPlugin.Shared;
using Sza.SolutionExportPlugin.Shared.Enum;
using XrmToolBox.Extensibility;
using XrmToolBox.Extensibility.Args;
using XrmToolBox.Extensibility.Interfaces;

namespace Sza.SolutionExportPlugin
{
    public  class ExportSolutions : PluginControlBase, IStatusBarMessager, IGitHubPlugin
    {
        private IContainer components;
        private ToolStrip toolStripMenu;
        private ToolStripButton tsbCancel;
        private ToolStripButton btnGetSolutions;
        private ToolStripButton ExportSelected;
        private DataGridView dataGridView1;
        private ComboBox comboBox1;
        private System.Windows.Forms.Label label1;
        private CheckBox managedCheckBox;
        private CheckBox unmanagedCheckBox;
        private FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.Label label3;
        private Button button1;
        private TextBox outputDirBox;
        private System.Windows.Forms.Label label4;
        private CheckedListBox checkedListBox1;
        private SplitContainer splitContainer1;
        private DataGridViewTextBoxColumn Solution;
        private DataGridViewTextBoxColumn Version;
        private DataGridViewCheckBoxColumn Selector;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label2;
        private string version;

        public string RepositoryName => "XrmToolBoxPlugins";

        public string UserName => "saadzag";

        public event EventHandler<StatusBarMessageEventArgs> SendMessageToStatusBar;

        public ExportSolutions()
        {
            this.InitializeComponent();
        }

        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ExportSolutions));
            this.toolStripMenu = new System.Windows.Forms.ToolStrip();
            this.btnGetSolutions = new System.Windows.Forms.ToolStripButton();
            this.ExportSelected = new System.Windows.Forms.ToolStripButton();
            this.tsbCancel = new System.Windows.Forms.ToolStripButton();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.Solution = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Version = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Selector = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.managedCheckBox = new System.Windows.Forms.CheckBox();
            this.unmanagedCheckBox = new System.Windows.Forms.CheckBox();
            this.label2 = new System.Windows.Forms.Label();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.label3 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.outputDirBox = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.checkedListBox1 = new System.Windows.Forms.CheckedListBox();
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.label5 = new System.Windows.Forms.Label();
            this.toolStripMenu.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            this.SuspendLayout();
            // 
            // toolStripMenu
            // 
            this.toolStripMenu.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.toolStripMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.btnGetSolutions,
            this.ExportSelected,
            this.tsbCancel});
            this.toolStripMenu.Location = new System.Drawing.Point(0, 0);
            this.toolStripMenu.Name = "toolStripMenu";
            this.toolStripMenu.Size = new System.Drawing.Size(540, 25);
            this.toolStripMenu.TabIndex = 4;
            this.toolStripMenu.Text = "toolStrip1";
            // 
            // btnGetSolutions
            // 
            this.btnGetSolutions.AccessibleName = "fgh";
            this.btnGetSolutions.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.btnGetSolutions.Image = ((System.Drawing.Image)(resources.GetObject("btnGetSolutions.Image")));
            this.btnGetSolutions.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnGetSolutions.Name = "btnGetSolutions";
            this.btnGetSolutions.Size = new System.Drawing.Size(81, 22);
            this.btnGetSolutions.Text = "Get Solutions";
            this.btnGetSolutions.Click += new System.EventHandler(this.GetSolutions_Click);
            // 
            // ExportSelected
            // 
            this.ExportSelected.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.ExportSelected.Image = ((System.Drawing.Image)(resources.GetObject("ExportSelected.Image")));
            this.ExportSelected.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.ExportSelected.Name = "ExportSelected";
            this.ExportSelected.Size = new System.Drawing.Size(91, 22);
            this.ExportSelected.Text = "Export Selected";
            this.ExportSelected.Click += new System.EventHandler(this.ExportSelected_Click);
            // 
            // tsbCancel
            // 
            this.tsbCancel.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.tsbCancel.Name = "tsbCancel";
            this.tsbCancel.Size = new System.Drawing.Size(47, 22);
            this.tsbCancel.Text = "Cancel";
            this.tsbCancel.ToolTipText = "Cancel the current request";
            this.tsbCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Solution,
            this.Version,
            this.Selector});
            this.dataGridView1.Location = new System.Drawing.Point(3, 3);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowHeadersVisible = false;
            this.dataGridView1.RowHeadersWidth = 82;
            this.dataGridView1.Size = new System.Drawing.Size(270, 636);
            this.dataGridView1.TabIndex = 7;
            // 
            // Solution
            // 
            this.Solution.FillWeight = 83.68021F;
            this.Solution.HeaderText = "Solution";
            this.Solution.MinimumWidth = 10;
            this.Solution.Name = "Solution";
            this.Solution.ReadOnly = true;
            // 
            // Version
            // 
            this.Version.FillWeight = 83.68021F;
            this.Version.HeaderText = "Version";
            this.Version.MinimumWidth = 10;
            this.Version.Name = "Version";
            this.Version.ReadOnly = true;
            // 
            // Selector
            // 
            this.Selector.FalseValue = false;
            this.Selector.FillWeight = 42.6396F;
            this.Selector.HeaderText = "";
            this.Selector.MinimumWidth = 10;
            this.Selector.Name = "Selector";
            this.Selector.TrueValue = true;
            // 
            // comboBox1
            // 
            this.comboBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.comboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(24, 70);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(125, 21);
            this.comboBox1.TabIndex = 8;
            this.comboBox1.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(21, 44);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(75, 13);
            this.label1.TabIndex = 9;
            this.label1.Text = "Target version";
            // 
            // managedCheckBox
            // 
            this.managedCheckBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.managedCheckBox.AutoSize = true;
            this.managedCheckBox.Location = new System.Drawing.Point(24, 142);
            this.managedCheckBox.Name = "managedCheckBox";
            this.managedCheckBox.Size = new System.Drawing.Size(71, 17);
            this.managedCheckBox.TabIndex = 10;
            this.managedCheckBox.Text = "Managed";
            this.managedCheckBox.UseVisualStyleBackColor = true;
            // 
            // unmanagedCheckBox
            // 
            this.unmanagedCheckBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.unmanagedCheckBox.AutoSize = true;
            this.unmanagedCheckBox.Location = new System.Drawing.Point(24, 175);
            this.unmanagedCheckBox.Name = "unmanagedCheckBox";
            this.unmanagedCheckBox.Size = new System.Drawing.Size(84, 17);
            this.unmanagedCheckBox.TabIndex = 11;
            this.unmanagedCheckBox.Text = "Unmanaged";
            this.unmanagedCheckBox.UseVisualStyleBackColor = true;
            // 
            // label2
            // 
            this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(21, 118);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(79, 13);
            this.label2.TabIndex = 12;
            this.label2.Text = "Package type :";
            // 
            // label3
            // 
            this.label3.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(21, 225);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(88, 13);
            this.label3.TabIndex = 13;
            this.label3.Text = "Output directory :";
            // 
            // button1
            // 
            this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.button1.Location = new System.Drawing.Point(147, 242);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 14;
            this.button1.Text = "Browse";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // outputDirBox
            // 
            this.outputDirBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.outputDirBox.Location = new System.Drawing.Point(24, 245);
            this.outputDirBox.Name = "outputDirBox";
            this.outputDirBox.Size = new System.Drawing.Size(117, 20);
            this.outputDirBox.TabIndex = 15;
            // 
            // label4
            // 
            this.label4.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(21, 265);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(82, 13);
            this.label4.TabIndex = 16;
            this.label4.Text = "Export settings :";
            // 
            // checkedListBox1
            // 
            this.checkedListBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.checkedListBox1.FormattingEnabled = true;
            this.checkedListBox1.Items.AddRange(new object[] {
            "Auto-numbering",
            "Calendar",
            "Customization",
            "Email tracking",
            "General",
            "Marketing",
            "Outlook Synchronization",
            "Relationship Roles",
            "ISV Config",
            "Sales"});
            this.checkedListBox1.Location = new System.Drawing.Point(24, 313);
            this.checkedListBox1.MaximumSize = new System.Drawing.Size(170, 184);
            this.checkedListBox1.MinimumSize = new System.Drawing.Size(170, 184);
            this.checkedListBox1.Name = "checkedListBox1";
            this.checkedListBox1.Size = new System.Drawing.Size(170, 184);
            this.checkedListBox1.TabIndex = 17;
            // 
            // splitContainer1
            // 
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.Location = new System.Drawing.Point(0, 25);
            this.splitContainer1.Name = "splitContainer1";
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.label5);
            this.splitContainer1.Panel1.Controls.Add(this.checkedListBox1);
            this.splitContainer1.Panel1.Controls.Add(this.label4);
            this.splitContainer1.Panel1.Controls.Add(this.outputDirBox);
            this.splitContainer1.Panel1.Controls.Add(this.button1);
            this.splitContainer1.Panel1.Controls.Add(this.label3);
            this.splitContainer1.Panel1.Controls.Add(this.label2);
            this.splitContainer1.Panel1.Controls.Add(this.managedCheckBox);
            this.splitContainer1.Panel1.Controls.Add(this.unmanagedCheckBox);
            this.splitContainer1.Panel1.Controls.Add(this.label1);
            this.splitContainer1.Panel1.Controls.Add(this.comboBox1);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.dataGridView1);
            this.splitContainer1.Size = new System.Drawing.Size(540, 759);
            this.splitContainer1.SplitterDistance = 260;
            this.splitContainer1.TabIndex = 18;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.ForeColor = System.Drawing.Color.Red;
            this.label5.Location = new System.Drawing.Point(21, 16);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(0, 13);
            this.label5.TabIndex = 18;
            // 
            // ExportSolutions
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Inherit;
            this.AutoSize = true;
            this.Controls.Add(this.splitContainer1);
            this.Controls.Add(this.toolStripMenu);
            this.Name = "ExportSolutions";
            this.Size = new System.Drawing.Size(540, 784);
            this.toolStripMenu.ResumeLayout(false);
            this.toolStripMenu.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel1.PerformLayout();
            this.splitContainer1.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        private void SetTragets()
        {
            if(Utilities.Version(version) == "8.2")
            {
                this.comboBox1.Items.AddRange(Constants.targets82);
            }
            if (Utilities.Version(version) == "8.1")
            {
                this.comboBox1.Items.AddRange(Constants.targets81);
            }
            if (Utilities.Version(version) == "8.0")
            {
                this.comboBox1.Items.AddRange(Constants.targets80);
            }
            if (Utilities.Version(version).Contains("9."))
            {
                this.comboBox1.Items.AddRange(Constants.targets90);
            }
        }
        private void DisplayVersion()
        {
             version = PluginService.RetrieveVersionRequest(Service);
            this.label5.Text = "Your actual CRM version is " + version;
        }

        private void BtnCloseClick(object sender, EventArgs e)
        {
            CloseTool(); // PluginBaseControl method that notifies the XrmToolBox that the user wants to close the plugin
            // Override the ClosingPlugin method to allow for any plugin specific closing logic to be performed (saving configs, canceling close, etc...)
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            CancelWorker(); // PluginBaseControl method that calls the Background Workers CancelAsync method.

            MessageBox.Show("Cancelled");
        }

        private void GetSolutions_Click(object sender, EventArgs e)
        {
            DisplayVersion();
            SetTragets();
            dataGridView1.Rows.Clear();
            var solutions = PluginService.GetUnmanagedSolutions(Service);
            if (solutions.Entities.Count >0)
            {
                foreach (Entity item in solutions.Entities)
                {
                    dataGridView1.Rows.Add(new object[] {item.Attributes["uniquename"].ToString(), item.Attributes["version"].ToString(),false });
                }
            }
            else
            {
                MessageBox.Show("No Solution found");
            }
        }

        private void ExportSelected_Click(object sender, EventArgs e)
        {
            dataGridView1.Refresh();
            dataGridView1.CurrentCell = dataGridView1.Rows[1].Cells[1];
            //var linq = (from row in dataGridView1.Rows.Cast<DataGridViewRow>()
            //            from cell in row.Cells.Cast<DataGridViewCell>()
            //            select new {  }).ToList();
            if (string.IsNullOrEmpty(outputDirBox.Text))
            {
                MessageBox.Show("Please fill in the output directory");
                return;

            }
            else if(!Directory.Exists(outputDirBox.Text)){
                MessageBox.Show("Please fill in  correctly the output directory");
                return;
            }
            
            foreach (DataGridViewRow dataGridRow in dataGridView1.Rows)
            {
                var solution = dataGridRow.Cells["Solution"].Value;
                if (/*dataGridRow.Cells["Selector"].Value != null &&*/
                     (bool)dataGridRow.Cells["Selector"].Value == true)
                {
                    ProcessExporting(dataGridRow);
               
                }
            
                
                //else if (dataGridRow.Cells["Selector"].Value == null)
                //{
                    
                //}
            }
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
       
        }
        private void ProcessExporting(DataGridViewRow dataGridRow)
        {
            string targetVersionItem = comboBox1.SelectedItem.ToString();
            bool isManaged = managedCheckBox.Checked ? true : false;
            bool isUnManaged = unmanagedCheckBox.Checked ? true : false;
            var exportRequest = ExportSolutionMapping(dataGridRow);
            var outputDir = outputDirBox.Text;
            string error = "";
            WorkAsync(new WorkAsyncInfo
            {
                Message = "Exporting " + (string)dataGridRow.Cells["Solution"].Value,
                Work = (w, e) =>
                {
                    
                    if (isManaged == true && isUnManaged == true)
                    {
                        exportRequest.Managed = true;
                        PluginService.ExportSolution(Service, ConnectionDetail, exportRequest, (string)dataGridRow.Cells["Version"].Value,outputDir);
                        exportRequest.Managed = false;
                        error = PluginService.ExportSolution(Service, ConnectionDetail, exportRequest, (string)dataGridRow.Cells["Version"].Value, outputDir);
                    
                    }
                    else if (isManaged == false && isUnManaged == true)
                    {
                        exportRequest.Managed = false;
                        error = PluginService.ExportSolution(Service, ConnectionDetail, exportRequest, (string)dataGridRow.Cells["Version"].Value, outputDir);
                    }
                    else if(isManaged == true && isUnManaged == false) {
                        exportRequest.Managed = true;
                        error = PluginService.ExportSolution(Service,ConnectionDetail, exportRequest, (string)dataGridRow.Cells["Version"].Value, outputDir);
                    }
                    
                },
                ProgressChanged = e =>
                {
                    // If progress has to be notified to user, use the following method:
                    SetWorkingMessage("Message to display");
                },
                PostWorkCallBack = e =>
                {
                    MessageBox.Show(error);
                },
                AsyncArgument = null,
                IsCancelable = true,
                MessageWidth = 340,
                MessageHeight = 150
            });
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                outputDirBox.Clear();
                outputDirBox.Text = folderBrowserDialog1.SelectedPath;
            }
        }
        private ExportSolutionRequest ExportSolutionMapping(DataGridViewRow dataGridRow)
        {
            string targetVersionItem = comboBox1.SelectedItem.ToString();
            ExportSolutionRequest exportSolutionRequest = new ExportSolutionRequest();
            exportSolutionRequest.SolutionName = (string)dataGridRow.Cells["Solution"].Value;
           if(!targetVersionItem.Contains("9.")) exportSolutionRequest.TargetVersion = targetVersionItem;

           if (checkedListBox1.CheckedItems.Count == 0) return exportSolutionRequest;
           for (var x = 0; x <= checkedListBox1.CheckedItems.Count - 1; x++)
            {
                var checkedItem = checkedListBox1.CheckedItems[x].ToString();
                exportSolutionRequest.ExportAutoNumberingSettings = this.checkedListBox1.Items[0].Equals(checkedItem) ? true : false;
                exportSolutionRequest.ExportCalendarSettings = this.checkedListBox1.Items[1].Equals(checkedItem) ? true : false; 
                exportSolutionRequest.ExportCustomizationSettings = this.checkedListBox1.Items[2].Equals(checkedItem) ? true : false; 
                exportSolutionRequest.ExportEmailTrackingSettings = this.checkedListBox1.Items[3].Equals(checkedItem) ? true : false; 
                exportSolutionRequest.ExportGeneralSettings = this.checkedListBox1.Items[4].Equals(checkedItem) ? true : false; 
                exportSolutionRequest.ExportIsvConfig = this.checkedListBox1.Items[5].Equals(checkedItem) ? true : false; 
                exportSolutionRequest.ExportMarketingSettings = this.checkedListBox1.Items[6].Equals(checkedItem) ? true : false; 
                exportSolutionRequest.ExportOutlookSynchronizationSettings = this.checkedListBox1.Items[7].Equals(checkedItem) ? true : false;
                exportSolutionRequest.ExportRelationshipRoles = this.checkedListBox1.Items[8].Equals(checkedItem) ? true : false; 
                exportSolutionRequest.ExportSales = this.checkedListBox1.Items[9].Equals(checkedItem) ? true : false;
            }

            return exportSolutionRequest;
        }

    }
}

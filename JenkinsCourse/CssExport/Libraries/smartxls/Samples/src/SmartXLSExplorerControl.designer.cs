namespace SmartXLSExplorer
{
    partial class SmartXLSExplorerControl
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

        #region Component Designer generated code

        private void InitializeComponent()
        {
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.navigationUserControl1 = new SmartXLSExplorer.SampleNavigationControl();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.samplePanel = new System.Windows.Forms.WebBrowser();
            this.sampleContainer1 = new SmartXLSExplorer.SampleContainerControl();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            this.tableLayoutPanel1.SuspendLayout();
            this.SuspendLayout();

            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.FixedPanel = System.Windows.Forms.FixedPanel.Panel1;
            this.splitContainer1.Location = new System.Drawing.Point(0, 0);
            this.splitContainer1.Margin = new System.Windows.Forms.Padding(2);
            this.splitContainer1.Name = "splitContainer1";

            this.splitContainer1.Panel1.Controls.Add(this.navigationUserControl1);

            this.splitContainer1.Panel2.Controls.Add(this.tableLayoutPanel1);
            this.splitContainer1.Size = new System.Drawing.Size(484, 323);
            this.splitContainer1.SplitterDistance = 290;
            this.splitContainer1.SplitterWidth = 3;
            this.splitContainer1.TabIndex = 0;

            this.navigationUserControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.navigationUserControl1.Location = new System.Drawing.Point(0, 0);
            this.navigationUserControl1.Margin = new System.Windows.Forms.Padding(2);
            this.navigationUserControl1.Name = "navigationUserControl1";
            this.navigationUserControl1.Size = new System.Drawing.Size(290, 323);
            this.navigationUserControl1.TabIndex = 0;
            this.navigationUserControl1.SelectionChanged += new SmartXLSExplorer.SelectionChangedEventHandler(this.navigationControl1_SelectionChanged);

            this.tableLayoutPanel1.ColumnCount = 1;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Controls.Add(this.samplePanel, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.sampleContainer1, 0, 0);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Margin = new System.Windows.Forms.Padding(2);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 2;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(191, 323);
            this.tableLayoutPanel1.TabIndex = 0;

            this.samplePanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.samplePanel.Location = new System.Drawing.Point(3, 7);
            this.samplePanel.MinimumSize = new System.Drawing.Size(20, 20);
            this.samplePanel.Name = "sample panel";
            this.samplePanel.ScriptErrorsSuppressed = true;
            this.samplePanel.Size = new System.Drawing.Size(185, 313);
            this.samplePanel.TabIndex = 1;

            this.sampleContainer1.AutoSize = true;
            this.sampleContainer1.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.sampleContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.sampleContainer1.Location = new System.Drawing.Point(2, 2);
            this.sampleContainer1.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.sampleContainer1.Name = "sampleContainer1";
            this.sampleContainer1.Size = new System.Drawing.Size(187, 1);
            this.sampleContainer1.TabIndex = 0;

            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.splitContainer1);
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "SmartXLSExplorerControl";
            this.Size = new System.Drawing.Size(484, 323);
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            this.splitContainer1.ResumeLayout(false);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.WebBrowser samplePanel;
        private SampleNavigationControl navigationUserControl1;
        private SampleContainerControl sampleContainer1;
    }
}

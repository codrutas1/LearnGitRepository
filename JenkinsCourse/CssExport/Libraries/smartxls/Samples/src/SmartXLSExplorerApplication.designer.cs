namespace SmartXLSExplorer
{
    partial class SmartXLSExplorerApplication
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
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.smartXLSExplorerUserControl2 = new SmartXLSExplorer.SmartXLSExplorerControl();
            this.SuspendLayout();
            // 
            // statusStrip1
            // 
            this.statusStrip1.Location = new System.Drawing.Point(0, 549);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Padding = new System.Windows.Forms.Padding(1, 0, 10, 0);
            this.statusStrip1.Size = new System.Drawing.Size(792, 22);
            this.statusStrip1.TabIndex = 1;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // smartXLSExplorerUserControl2
            // 
            this.smartXLSExplorerUserControl2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.smartXLSExplorerUserControl2.Location = new System.Drawing.Point(0, 0);
            this.smartXLSExplorerUserControl2.Margin = new System.Windows.Forms.Padding(2);
            this.smartXLSExplorerUserControl2.MinimumSize = new System.Drawing.Size(375, 284);
            this.smartXLSExplorerUserControl2.Name = "smartXLSExplorerUserControl2";
            this.smartXLSExplorerUserControl2.Size = new System.Drawing.Size(792, 549);
            this.smartXLSExplorerUserControl2.TabIndex = 0;
            // 
            // SmartXLSExplorerApplication
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.ClientSize = new System.Drawing.Size(792, 571);
            this.Controls.Add(this.smartXLSExplorerUserControl2);
            this.Controls.Add(this.statusStrip1);
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "SmartXLSExplorerApplication";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "SmartXLS Explorer";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private SmartXLSExplorer.SmartXLSExplorerControl smartXLSExplorerUserControl2;
        private System.Windows.Forms.StatusStrip statusStrip1;
    }
}


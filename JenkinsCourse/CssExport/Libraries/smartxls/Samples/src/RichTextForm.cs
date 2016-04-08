using System.Windows.Forms;

namespace SmartXLSExplorer
{
    public class RichTextForm : Form
    {
        public RichTextBox richTextBox1;
        private System.ComponentModel.Container components = null;

        public RichTextForm()
        {
            InitializeComponent();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (components != null)
                {
                    components.Dispose();
                }
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.richTextBox1 = new RichTextBox();
            this.SuspendLayout();
            this.richTextBox1.Location = new System.Drawing.Point(24, 24);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.Size = new System.Drawing.Size(160, 70);
            this.richTextBox1.TabIndex = 1;
            this.richTextBox1.Text = "";
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(200, 100);
            this.Controls.Add(richTextBox1);
            this.Name = "Form1";
            this.Text = "SmartXLS Samples";
            this.ResumeLayout(false);

        }
    }
}

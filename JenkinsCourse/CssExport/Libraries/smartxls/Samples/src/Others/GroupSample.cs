using System;
using System.Drawing;
using System.Windows.Forms;
using SmartXLS;

namespace SmartXLSExplorer
{
    public class GroupSample : BaseFileViewControl, ISampleControl
    {
        public GroupSample()
        {
            InitializeComponent();
            base.buttonViewFile.Text = "Read";
        }

        public override void buttonView_Click(object sender, EventArgs e)
        {
            WorkBook book1 = new WorkBook();
            WorkBook book2 = new WorkBook();
            try
            {
                book2.setNumber(0, 0, 458);
                book1.setWorkbookName("wb1");
                book2.setWorkbookName("wb2");
                book1.setGroup("group");
                book2.setGroup("group");
                book1.setFormula(0, 1, 1, "SUM('[wb2]Sheet1'!$A$1:$D$4)");
                double result = book1.getNumber(1, 1);
                MessageBox.Show("SUM('[wb2]Sheet1'!$A$1:$D$4) = " + Convert.ToString(result), "Result", MessageBoxButtons.OK, MessageBoxIcon.Information);
                //Console.Out.WriteLine("Result:" + result);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine(ex);
            }
        }

        string ISampleControl.CodePath
        {
            get
            {
                return "src\\Others\\GroupSample.cs";
            }
        }
    }
}

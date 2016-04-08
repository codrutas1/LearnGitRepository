using System;
using System.Windows.Forms;
using SmartXLS;

namespace SmartXLSExplorer
{
    public partial class ChartPreserveSample : BaseFileViewControl, ISampleControl
    {
        public ChartPreserveSample()
            : base()
        {
            InitializeComponent();
        }

        public override void buttonView_Click(object sender, EventArgs e)
        {
            try
            {
                WorkBook book = new WorkBook();

                book.read(@"..\template\ChartPropertiesTemplate.xls");

                //Saving the workbook to disk.
                book.write(".\\chartProp.xls");
            }
            catch (Exception ex)
            {
                System.Console.Out.WriteLine(ex);
            }

            //Message box confirmation to view the created spreadsheet.
            if (MessageBox.Show("Do you want to view the workbook?", "Workbook has been created",
                MessageBoxButtons.YesNo, MessageBoxIcon.Information)
                == DialogResult.Yes)
            {
                //Launching the Excel file using the default Application.[MS Excel Or Free ExcelViewer]
                System.Diagnostics.Process.Start(".\\chartProp.xls");
            }
            else
            {
            }
        }

        string ISampleControl.CodePath
        {
            get
            {
                return "src\\Charts\\ChartPreserveSample.cs";
            }
        }
    }
}

using System;
using System.Windows.Forms;
using SmartXLS;

namespace SmartXLSExplorer
{
    public class ChartEditSample : BaseFileViewControl, ISampleControl
    {
        public ChartEditSample()
        {
            InitializeComponent();
        }

        public override void buttonView_Click(object sender, EventArgs e)
        {
            try
            {
                WorkBook workBook = new WorkBook();

                //read in
                workBook.read("..\\template\\chartTemplate.xls");

                //get chartshape from sheet 1
                ChartShape chartShape = workBook.getChart(0);

                chartShape.ChartType = ChartShape.Bar;
                chartShape.Title = "Chart 1";
                //change 3D chart to 2D
                chartShape.set3Dimensional(false);

                //select sheet 2
                workBook.Sheet = 1;
                //get chartshape in the sheet
                chartShape = workBook.getChart(0);
                //change chart type to step
                chartShape.ChartType = ChartShape.Step;
                //set axis title
                chartShape.setAxisTitle(ChartShape.XAxis, 0, "X-axis data");
                chartShape.setAxisTitle(ChartShape.YAxis, 0, "Y-axis data");
                chartShape.Title = "Chart 2";
                //change chart to 3D
                chartShape.set3Dimensional(true);

                //write out
                workBook.write("./ChartEdit.xls");
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
                System.Diagnostics.Process.Start(".\\ChartEdit.xls");
            }
        }

        string ISampleControl.CodePath
        {
            get
            {
                return "src\\Charts\\ChartEditSample.cs";
            }
        }
    }
}

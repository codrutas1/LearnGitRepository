using System;
using System.Drawing;
using System.Windows.Forms;
using SmartXLS;


namespace SmartXLSExplorer
{
    public class ConditionFormatSample : BaseFileViewControl, ISampleControl
    {

        public ConditionFormatSample()
        {
            InitializeComponent();
        }

        public override void buttonView_Click(object sender, EventArgs e)
        {
            WorkBook workBook = new WorkBook();
            try
            {
                ConditionFormat[] condfmt = new ConditionFormat[3];
                condfmt[0] = workBook.CreateConditionFormat();
                condfmt[1] = workBook.CreateConditionFormat();
                condfmt[2] = workBook.CreateConditionFormat();

                // Condition #1
                RangeStyle cf = condfmt[0].RangeStyle;
                condfmt[0].Type = ConditionFormat.eTypeFormula;
                condfmt[0].setFormula1("and(iseven(row()), $D1 > 1000)", 0, 0);
                cf.FontColor = 0x00ff00;
                cf.Pattern = RangeStyle.PatternSolid;
                cf.PatternFG = 0xcc99ff;
                condfmt[0].RangeStyle = cf;

                // Condition #2
                condfmt[1].Type = ConditionFormat.eTypeFormula;
                condfmt[1].setFormula1("iseven($A1)", 0, 0);
                cf.FontColor = 0xffffff;
                condfmt[1].RangeStyle = cf;

                // Condition #3
                condfmt[2].Type = ConditionFormat.eTypeCell;
                condfmt[2].setFormula1("500", 0, 0);
                condfmt[2].Operator = ConditionFormat.eOperatorGreaterThan;
                cf = condfmt[2].RangeStyle;
                cf.FontColor = 0xff0000;
                condfmt[2].RangeStyle = cf;

                // Select the range and apply conditional formatting
                workBook.setSelection(0, 0, 39, 3);
                workBook.ConditionalFormats = condfmt;

                workBook.write("./Sample.xls");

                //Message box confirmation to view the created spreadsheet.
                if (MessageBox.Show("Do you want to view the workbook?", "Workbook has been created",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Information)
                    == DialogResult.Yes)
                {
                    //Launching the Excel file using the default Application.[MS Excel Or Free ExcelViewer]
                    System.Diagnostics.Process.Start("Sample.xls");
                }
            }
            catch (System.Exception ex)
            {
               Console.WriteLine(ex.Message);
            }
        }

        string ISampleControl.CodePath
        {
            get
            {
                return "src\\Formatting\\ConditionFormatSample.cs";
            }
        }
    }
}

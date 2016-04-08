using System;
using System.Drawing;
using System.Windows.Forms;
using SmartXLS;

namespace SmartXLSExplorer
{
    public class DataValidationWriteSample : BaseFileViewControl, ISampleControl
    {
        public DataValidationWriteSample()
        {
            InitializeComponent();
        }

        public override void buttonView_Click(object sender, EventArgs e)
        {
            try
            {
                WorkBook workBook = new WorkBook();

                workBook.setText(0, 1, "Apple");
                workBook.setText(0, 2, "Orange");
                workBook.setText(0, 3, "Banana");
                workBook.setText(6, 3, "the input value in cell C7 must be between 0 to 10");

                DataValidation dataValidation = workBook.CreateDataValidation();
                dataValidation.Type = DataValidation.eUser;
                dataValidation.Formula1 = "\"dddd\x0000gggg\x0000hhh\"";
                workBook.setSelection("A1:A5");
                workBook.DataValidation = dataValidation;

                dataValidation = workBook.CreateDataValidation();
                dataValidation.Type = DataValidation.eUser;
                dataValidation.Formula1 = "$B$1:$D$1";
                workBook.setSelection("B1:D5");
                workBook.DataValidation = dataValidation;

                dataValidation = workBook.CreateDataValidation();
                dataValidation.Type = DataValidation.eInteger;
                dataValidation.Operator = DataValidation.eBetween;
                dataValidation.Formula1 = "0";
                dataValidation.Formula2 = "10";
                workBook.setSelection(6, 2, 6, 2);//select C7
                workBook.DataValidation = dataValidation;

                workBook.write(".\\datavalidation.xls");

                //Message box confirmation to view the created spreadsheet.
                if (MessageBox.Show("Do you want to view the workbook?", "Workbook has been created",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Information)
                    == DialogResult.Yes)
                {
                    //Launching the Excel file using the default Application.[MS Excel Or Free ExcelViewer]
                    System.Diagnostics.Process.Start(".\\datavalidation.xls");
                }
                else
                {
                }
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
                return "src\\DataValidation\\DataValidationWriteSample.cs";
            }
        }
    }
}

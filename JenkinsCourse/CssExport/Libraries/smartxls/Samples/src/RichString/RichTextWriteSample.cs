using System;
using System.Drawing;
using System.Windows.Forms;
using SmartXLS;

namespace SmartXLSExplorer
{
    public class RichTextWriteSample : BaseFileViewControl, ISampleControl
    {
        public RichTextWriteSample()
        {
            InitializeComponent();
        }

        public override void buttonView_Click(object sender, EventArgs e)
        {
            try
            {
                WorkBook workBook = new WorkBook();

                //set data
                workBook.setText(0, 0, "Hello, you are welcome!");

                //text orientation
                RangeStyle rangeStyle = workBook.getRangeStyle();
                rangeStyle.Orientation = (short)45;
                workBook.setRangeStyle(rangeStyle);

                //multi text selection format
                workBook.setTextSelection(0, 6);
                rangeStyle = workBook.getRangeStyle();
                rangeStyle.FontBold = true;
                workBook.setRangeStyle(rangeStyle);

                workBook.setTextSelection(7, 10);
                rangeStyle = workBook.getRangeStyle();
                rangeStyle.FontItalic = true;
                workBook.setRangeStyle(rangeStyle);

                workBook.setTextSelection(11, 14);
                rangeStyle = workBook.getRangeStyle();
                rangeStyle.FontUnderline = RangeStyle.UnderlineSingle;
                workBook.setRangeStyle(rangeStyle);

                workBook.setTextSelection(15, 22);
                rangeStyle = workBook.getRangeStyle();
                rangeStyle.FontSize = 14 * 20;
                workBook.setRangeStyle(rangeStyle);

                workBook.write("./TextFormatting.xls");


                //Message box confirmation to view the created spreadsheet.
                if (MessageBox.Show("Do you want to view the workbook?", "Workbook has been created",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Information)
                    == DialogResult.Yes)
                {
                    //Launching the Excel file using the default Application.[MS Excel Or Free ExcelViewer]
                    System.Diagnostics.Process.Start(".\\TextFormatting.xls");
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
                return "src\\RichString\\RichTextWriteSample.cs";
            }
        }
    }
}

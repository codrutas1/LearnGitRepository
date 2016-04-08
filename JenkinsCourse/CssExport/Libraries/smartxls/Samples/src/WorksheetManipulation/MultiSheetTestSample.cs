using System;
using System.IO;
using System.Windows.Forms;
using SmartXLS;

namespace SmartXLSExplorer
{
    public class MultiSheetTestSample : BaseFileViewControl, ISampleControl
    {
        public MultiSheetTestSample()
        {
            InitializeComponent();
        }

        public override void buttonView_Click(object sender, EventArgs e)
        {
            WorkBook m_book = null;
            try
            {
                m_book = new WorkBook();
                m_book.setSheetName(0, "sheet1");
                m_book.setText(1, 1, "sheet1");

                m_book.insertSheets(1, 1);
                m_book.setSheetName(1, "sheet2");
                m_book.Sheet = 1;
                m_book.setText(2, 2, "sheet2");
                m_book.insertSheets(2, 1);
                m_book.setSheetName(2, "sheet3");
                m_book.Sheet = 2;
                m_book.setText(3, 3, "sheet3");
                m_book.setText(1, 3, "sheet3");

                m_book.Sheet = 0;
                m_book.write(".\\Sample.xls");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine(ex);
            }
 
            //Message box confirmation to view the created spreadsheet.
            if (MessageBox.Show("Do you want to view the workbook?", "Workbook has been created",
                MessageBoxButtons.YesNo, MessageBoxIcon.Information)
                == DialogResult.Yes)
            {
                //Launching the Excel file using the default Application.[MS Excel Or Free ExcelViewer]
                System.Diagnostics.Process.Start("Sample.xls");
            }
            else
            {
            }
        }

        string ISampleControl.CodePath
        {
            get
            {
                return "src\\WorksheetManipulation\\MultiSheetTestSample.cs";
            }
        }
    }
}

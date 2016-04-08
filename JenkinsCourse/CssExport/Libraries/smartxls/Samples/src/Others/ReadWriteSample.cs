using System;
using System.Drawing;
using System.Windows.Forms;
using SmartXLS;

namespace SmartXLSExplorer
{
    public partial class ReadWriteSample : BaseFileViewControl, ISampleControl
    {
        public ReadWriteSample()
            : base()
        {
            InitializeComponent();
        }

        public override void buttonView_Click(object sender, EventArgs e)
        {
            WorkBook m_book = new WorkBook();

            int row = 0;
            int col = 0;
            try
            {
                m_book.setSheetName(0, "First tab");
                m_book.insertSheets(1, 1);
                m_book.setSheetName(1, "Second tab");
                m_book.Sheet = 1;

                //m_book.AutoRecalc = false;
                m_book.setText(65535, 0, "");
                for (row = 0; row < 5536; row++)
                {
                    for (col = 0; col < 5; col++)
                    {
                        String text = System.Convert.ToString(row) + System.Convert.ToString(col);
                        if (col % 2 == 1)
                        {
                            if (row == 5535)
                            {
                                String start = m_book.formatRCNr(0, col, false);
                                String end = m_book.formatRCNr(row - 1, col, false);
                                String formula = "SUM(" + start + ":" + end + ")";
                                Console.Out.WriteLine(formula);
                                m_book.setFormula(row, col, formula);
                            }
                            else
                                m_book.setNumber(row, col, Double.Parse(text));
                        }
                        else
                            m_book.setText(row, col, text);
                    }
                    m_book.setSelection(row, 0, row, col - 1);
                    RangeStyle cfmt = m_book.getRangeStyle();
                    cfmt.Pattern = (short)1;
                    cfmt.PatternBG = m_book.getPaletteEntry(row % 55);
                }
                m_book.recalc();
                m_book.write(".\\Sample.xls");
            }
            catch (Exception ex)
            {
                System.Console.Out.WriteLine(row);
                System.Console.Out.WriteLine(ex.Message);
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
                return "src\\Others\\ReadWriteSample.cs";
            }
        }
    }
}

using System;
using System.Drawing;
using System.Windows.Forms;
using SmartXLS;

namespace SmartXLSExplorer
{
    public partial class PivotSample : BaseFileViewControl, ISampleControl
    {
        public PivotSample()
            : base()
        {
            InitializeComponent();
        }

        public override void buttonView_Click(object sender, EventArgs e)
        {
            WorkBook m_book = new WorkBook();

            try
            {
                m_book.read("..\\template\\PivotTable.xls");

                BookPivotRangeModel model = m_book.getPivotModel();
                //Sets the source range that should be used for the PivotRange
                model.setList("A1:D27");
                //Defines the location of the PivotRange.
                model.setLocation(0, 17, 5);

                //make the cell active(to make the pivotRange selected)
                m_book.setActiveCell(17, 5);
                //get the currently selected PivotRange
                BookPivotRange pivotRange = model.getActivePivotRange();
                //refresh the pivot table from the data source
                model.refreshRange(pivotRange);

                //get the Area object associated with the PivotRange.
                RangeArea rangeArea = pivotRange.getRangeArea();

                Console.WriteLine("PivotRange Scope:" + rangeArea.ToString());
                //pivotRange.addFormulaField("double amount:", "=Amount*2");

                BookPivotArea rowArea = pivotRange.getArea(BookPivotRange.row);
                BookPivotArea columnArea = pivotRange.getArea(BookPivotRange.column);
                BookPivotArea dataArea = pivotRange.getArea(BookPivotRange.data);
                BookPivotArea pageArea = pivotRange.getArea(BookPivotRange.page);

                BookPivotField rowField = pivotRange.getField("Who");
                rowArea.addField(rowField);
                BookPivotField dataField = pivotRange.getField("Amount");
                //BookPivotField dataField = pivotRange.getField("double amount:");
                dataArea.addField(dataField);
                BookPivotField columnField = pivotRange.getField("What");
                columnArea.addField(columnField);
                BookPivotField pageField = pivotRange.getField("Week");
                pageArea.addField(pageField);

                m_book.writeXLSX(".\\PivotTable.xlsx");
            }
            catch (Exception ex)
            {
                System.Console.Out.WriteLine(ex.Message);
            }

            //Message box confirmation to view the created spreadsheet.
            if (MessageBox.Show("Do you want to view the workbook?", "Workbook has been created",
                MessageBoxButtons.YesNo, MessageBoxIcon.Information)
                == DialogResult.Yes) 
            {
                //Launching the Excel file using the default Application.[MS Excel Or Free ExcelViewer]
                System.Diagnostics.Process.Start("PivotTable.xlsx");
            }
            else
            {
            }
        }

        string ISampleControl.CodePath
        {
            get
            {
                return "src\\Others\\PivotSample.cs";
            }
        }
    }
}

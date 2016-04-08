using System;
using System.Drawing;
using System.Windows.Forms;
using SmartXLS;


namespace SmartXLSExplorer
{
    public class CellFormatSample : BaseFileViewControl, ISampleControl
    {
        int StartRow, StartCol, EndRow, EndCol;
        string StartRange, eRange, hdrRange, colRange, ftrRange, bodyRange;
        WorkBook m_WorkBook;
        RangeStyle m_RangeStyle;
    
        public CellFormatSample()
        {
            InitializeComponent();
        }

        public override void buttonView_Click(object sender, EventArgs e)
        {
            m_WorkBook = new WorkBook();
            m_WorkBook.NumSheets = 14;
            m_WorkBook.setSheetName(0, "Simple Format");
            m_WorkBook.setSheetName(1, "Classic1 Format");
            m_WorkBook.setSheetName(2, "Classic2 Format");
            m_WorkBook.setSheetName(3, "Classic3 Format");
            m_WorkBook.setSheetName(4, "Accounting1 Format");
            m_WorkBook.setSheetName(5, "Accounting2 Format");
            m_WorkBook.setSheetName(6, "Accounting3 Format");
            m_WorkBook.setSheetName(7, "Effects3D1 Format");
            m_WorkBook.setSheetName(8, "Colorful1 Format");
            m_WorkBook.setSheetName(9, "Colorful2 Format");
            m_WorkBook.setSheetName(10, "Colorful3 Format");
            m_WorkBook.setSheetName(11, "List1 Format");
            m_WorkBook.setSheetName(12, "List2 Format");
            m_WorkBook.setSheetName(13, "List3 Format");

            m_WorkBook.Sheet = 0;
            setData();
            simpleFormat();

            m_WorkBook.Sheet = 1;
            setData();
            Classic1();

            m_WorkBook.Sheet = 2;
            setData();
            Classic2();

            m_WorkBook.Sheet = 3;
            setData();
            Classic3();

            m_WorkBook.Sheet = 4;
            setData();
            Accounting1();

            m_WorkBook.Sheet = 5;
            setData();
            Accounting2();

            m_WorkBook.Sheet = 6;
            setData();
            Accounting3();

            m_WorkBook.Sheet = 7;
            setData();
            Effects3D1();

            m_WorkBook.Sheet = 8;
            setData();
            Colorful1();

            m_WorkBook.Sheet = 9;
            setData();
            Colorful2();

            m_WorkBook.Sheet = 10;
            setData();
            Colorful3();

            m_WorkBook.Sheet = 11;
            setData();
            List1();

            m_WorkBook.Sheet = 12;
            setData();
            List2();

            m_WorkBook.Sheet = 13;
            setData();
            List3();

            m_WorkBook.write(".\\RangeStyle.xls");

            //Message box confirmation to view the created spreadsheet.
            if (MessageBox.Show("Do you want to view the workbook?", "Workbook has been created",
                MessageBoxButtons.YesNo, MessageBoxIcon.Information)
                == DialogResult.Yes)
            {
                //Launching the Excel file using the default Application.[MS Excel Or Free ExcelViewer]
                System.Diagnostics.Process.Start("RangeStyle.xls");
            }
            else
            {
            }
        }

        private void setData()
        {
            m_WorkBook.setText(1, 2, "Jan");
            m_WorkBook.setText(1, 3, "Feb");
            m_WorkBook.setText(1, 4, "Mar");
            m_WorkBook.setText(1, 5, "Apr");
            m_WorkBook.setText(2, 1, "Bananas");
            m_WorkBook.setText(3, 1, "Papaya");
            m_WorkBook.setText(4, 1, "Mango");
            m_WorkBook.setText(5, 1, "Lilikoi");
            m_WorkBook.setText(6, 1, "Comfrey");
            m_WorkBook.setText(7, 1, "Total");
            m_WorkBook.setFormula(2, 2, "RAND()*100");
            m_WorkBook.setSelection(2, 2, 2, 5);
            m_WorkBook.editCopyRight();
            m_WorkBook.setSelection(2, 2, 6, 5);
            m_WorkBook.editCopyDown();
            m_WorkBook.setFormula(7, 2, "SUM(C3:C7)");
            m_WorkBook.setSelection("C8:F8");
            m_WorkBook.editCopyRight();

            StartRow = 1;
            StartCol = 1;
            EndRow = 7;
            EndCol = 5;
            StartRange = m_WorkBook.formatRCNr(StartRow, StartCol, false);
            eRange = m_WorkBook.formatRCNr(StartRow, EndCol, false);
            hdrRange = StartRange + ":" + eRange;

            eRange = m_WorkBook.formatRCNr(EndRow, StartCol, false);
            colRange = StartRange + ":" + eRange;

            StartRange = m_WorkBook.formatRCNr(EndRow, StartCol, false);
            eRange = m_WorkBook.formatRCNr(EndRow, EndCol, false);
            ftrRange = StartRange + ":" + eRange;

            StartRange = m_WorkBook.formatRCNr(StartRow + 1, StartCol + 1, false);
            eRange = m_WorkBook.formatRCNr(EndRow - 1, EndCol, false);
            bodyRange = StartRange + ":" + eRange;

            m_RangeStyle = m_WorkBook.getRangeStyle();
        }

        private void simpleFormat()
        {
            m_WorkBook.setSelection(colRange);
            AdjustFont(Color.Black.ToArgb(), true, false, false);
            m_WorkBook.setRangeStyle(m_RangeStyle);

            m_WorkBook.setSelection(ftrRange);
            m_RangeStyle.BottomBorder = RangeStyle.BorderMedium;
            m_RangeStyle.TopBorder = RangeStyle.BorderThin;
            m_WorkBook.setRangeStyle(m_RangeStyle, EndRow, StartCol, EndRow, EndCol);

            m_WorkBook.setSelection(hdrRange);
            m_RangeStyle.TopBorder = RangeStyle.BorderMedium;
            m_RangeStyle.BottomBorder = RangeStyle.BorderMedium;
            m_RangeStyle.HorizontalAlignment = RangeStyle.HorizontalAlignmentRight;
            m_RangeStyle.VerticalAlignment = RangeStyle.VerticalAlignmentBottom;
            m_WorkBook.setRangeStyle(m_RangeStyle);

            m_WorkBook.setSelection(ftrRange);
            m_RangeStyle.BottomBorder = RangeStyle.BorderMedium;
            m_RangeStyle.TopBorder = RangeStyle.BorderThin;
            m_WorkBook.setRangeStyle(m_RangeStyle, EndRow, StartCol, EndRow, EndCol);
        }

        private void Classic1()
        {
            m_WorkBook.setSelection(colRange);
            m_RangeStyle.TopBorder = RangeStyle.BorderNone;
            m_RangeStyle.BottomBorder = RangeStyle.BorderNone;
            m_RangeStyle.RightBorder = RangeStyle.BorderThin;
            AdjustFont(Color.Black.ToArgb(), true, false, false);
            m_WorkBook.setRangeStyle(m_RangeStyle);

            m_WorkBook.setSelection(hdrRange);
            m_RangeStyle.VerticalInsideBorder = RangeStyle.BorderNone;
            m_RangeStyle.TopBorder = RangeStyle.BorderMedium;
            m_RangeStyle.BottomBorder = RangeStyle.BorderThin;
            AdjustFont(System.Drawing.Color.Black.ToArgb(), false, true, false);
            m_RangeStyle.HorizontalAlignment = RangeStyle.HorizontalAlignmentRight;
            m_RangeStyle.VerticalAlignment = RangeStyle.VerticalAlignmentBottom;
            m_WorkBook.setRangeStyle(m_RangeStyle);

            m_WorkBook.setSelection(ftrRange);
            m_RangeStyle.TopBorder = RangeStyle.BorderThin;
            m_RangeStyle.BottomBorder = RangeStyle.BorderMedium;
            m_WorkBook.setRangeStyle(m_RangeStyle);
        }

        private void Classic2()
        {
            m_WorkBook.setSelection(colRange);
            short nPattern;
            AdjustFont(Color.Black.ToArgb(), true, false, false);
            m_WorkBook.setRangeStyle(m_RangeStyle);

            m_WorkBook.setSelection(hdrRange);
            m_RangeStyle.TopBorder = RangeStyle.BorderMedium;
            m_RangeStyle.BottomBorder = RangeStyle.BorderThin;
            AdjustFont(Color.Black.ToArgb(), false, false, false);
            nPattern = 1;
            m_RangeStyle.Pattern = nPattern;
            m_RangeStyle.PatternFG = Color.Magenta.ToArgb();
            m_RangeStyle.PatternBG = Color.Magenta.ToArgb();
            m_RangeStyle.HorizontalAlignment = RangeStyle.HorizontalAlignmentRight;
            m_RangeStyle.VerticalAlignment = RangeStyle.VerticalAlignmentBottom;
            m_WorkBook.setRangeStyle(m_RangeStyle);

            m_WorkBook.setSelection(ftrRange);
            m_RangeStyle.BottomBorder = RangeStyle.BorderMedium;
            m_RangeStyle.TopBorder = RangeStyle.BorderThin;
            m_WorkBook.setRangeStyle(m_RangeStyle);
        }

        private void Classic3()
        {
            m_WorkBook.setSelection(hdrRange);
            m_RangeStyle.TopBorder = RangeStyle.BorderMedium;
            m_RangeStyle.BottomBorder = RangeStyle.BorderMedium;
            AdjustFont(Color.White.ToArgb(), true, true, false);
            AlignRight();
            SetSolidPattern(m_WorkBook.getPaletteEntry(11), Color.Black.ToArgb());
            m_WorkBook.setRangeStyle(m_RangeStyle);

            m_WorkBook.setSelection(ftrRange);
            m_RangeStyle.TopBorder = RangeStyle.BorderMedium;
            m_RangeStyle.BottomBorder = RangeStyle.BorderMedium;
            SetSolidPattern(m_WorkBook.getPaletteEntry(15), Color.Black.ToArgb());
            m_WorkBook.setRangeStyle(m_RangeStyle);

            m_WorkBook.setSelection(bodyRange);
            SetSolidPattern(m_WorkBook.getPaletteEntry(15), Color.Black.ToArgb());
            m_WorkBook.setRangeStyle(m_RangeStyle);

            m_WorkBook.setSelection(colRange);
            SetSolidPattern(m_WorkBook.getPaletteEntry(15), Color.Black.ToArgb());
            m_WorkBook.setRangeStyle(m_RangeStyle);
        }

        private void Accounting1()
        {
            m_WorkBook.setSelection(hdrRange);
            System.String numberFormat;
            m_RangeStyle.TopBorder = RangeStyle.BorderThin;
            m_RangeStyle.BottomBorder = RangeStyle.BorderThin;
            AdjustFont(Color.Magenta.ToArgb(), true, true, false);
            AlignRight();
            m_WorkBook.setRangeStyle(m_RangeStyle);

            m_WorkBook.setSelection(bodyRange);
            m_RangeStyle.BottomBorder = RangeStyle.BorderNone;
            numberFormat = "$ #,##0.00_);(#,##0.00)";
            m_RangeStyle.CustomFormat = numberFormat;
            m_WorkBook.setRangeStyle(m_RangeStyle);

            m_WorkBook.setSelection(ftrRange);
            m_RangeStyle.BottomBorder = RangeStyle.BorderDouble;
            m_RangeStyle.CustomFormat = numberFormat;
            m_WorkBook.setRangeStyle(m_RangeStyle);
        }

        private void Accounting2()
        {
            System.String numberFormat;

            m_WorkBook.setSelection(hdrRange);
            m_RangeStyle.TopBorder = RangeStyle.BorderThick;
            m_RangeStyle.TopBorderColor = Color.LightGray.ToArgb();
            m_RangeStyle.BottomBorder = RangeStyle.BorderThin;
            m_RangeStyle.BottomBorderColor = Color.LightGray.ToArgb();
            AlignRight();
            m_WorkBook.setRangeStyle(m_RangeStyle);

            m_WorkBook.setSelection(bodyRange);
            m_RangeStyle.BottomBorder = RangeStyle.BorderNone;
            numberFormat = "$ #,##0.00_);(#,##0.00)";
            m_RangeStyle.CustomFormat = numberFormat;
            m_WorkBook.setRangeStyle(m_RangeStyle);

            m_WorkBook.setSelection(ftrRange);
            m_RangeStyle.BottomBorder = RangeStyle.BorderThick;
            m_RangeStyle.BottomBorderColor = Color.LightGray.ToArgb();
            m_RangeStyle.TopBorder = RangeStyle.BorderThin;
            m_RangeStyle.TopBorderColor = Color.LightGray.ToArgb();
            numberFormat = "$ #,##0.00_);(#,##0.00)";
            m_RangeStyle.CustomFormat = numberFormat;
            m_WorkBook.setRangeStyle(m_RangeStyle);
        }

        private void Accounting3()
        {
            System.String numberFormat;
            m_WorkBook.setSelection(colRange);
            AdjustFont(Color.Black.ToArgb(), false, true, false);
            m_WorkBook.setRangeStyle(m_RangeStyle);

            m_WorkBook.setSelection(bodyRange);
            numberFormat = "#,##0.00_);(#,##0.00)";
            m_RangeStyle.CustomFormat = numberFormat;
            m_WorkBook.setRangeStyle(m_RangeStyle);

            m_WorkBook.setSelection(bodyRange);
            numberFormat = "$ #,##0.00_);(#,##0.00)";
            m_RangeStyle.TopBorder = RangeStyle.BorderNone;
            m_RangeStyle.BottomBorder = RangeStyle.BorderNone;
            m_RangeStyle.CustomFormat = numberFormat;
            m_WorkBook.setRangeStyle(m_RangeStyle);

            numberFormat = "$ #,##0.00_);(#,##0.00)";
            m_RangeStyle.CustomFormat = numberFormat;
            m_WorkBook.setRangeStyle(m_RangeStyle);

            m_WorkBook.setSelection(bodyRange);
            m_RangeStyle.TopBorder = RangeStyle.BorderThin;
            m_RangeStyle.BottomBorder = RangeStyle.BorderDouble;
            m_WorkBook.setRangeStyle(m_RangeStyle);

            m_WorkBook.setSelection(bodyRange);
            m_RangeStyle.TopBorder = RangeStyle.BorderNone;
            m_RangeStyle.BottomBorder = RangeStyle.BorderMedium;
            m_RangeStyle.BottomBorderColor = Color.Green.ToArgb();
            AdjustFont(m_WorkBook.getPaletteEntry(16), false, true, false);
            AlignRight();
            m_WorkBook.setRangeStyle(m_RangeStyle);
        }

        private void Effects3D1()
        {
            m_WorkBook.setSelection(1, 1, 7, 5);
            SetSolidPattern(Color.LightGray.ToArgb(), 0);

            Set3DBorder(2, 2, 6, 5, Color.LightGray.ToArgb(), Color.DarkGray.ToArgb(), Color.LightGray.ToArgb());
            m_WorkBook.setRangeStyle(m_RangeStyle);

            m_WorkBook.setSelection(hdrRange);
            AdjustFont(Color.Magenta.ToArgb(), true, false, false);
            AlignCenter();
            m_WorkBook.setRangeStyle(m_RangeStyle);

            m_WorkBook.setSelection(colRange);
            Set3DBorder(1, 1, 7, 1, Color.LightGray.ToArgb(), Color.LightGray.ToArgb(), Color.DarkGray.ToArgb());
            AdjustFont(Color.Black.ToArgb(), true, false, false);
            m_WorkBook.setRangeStyle(m_RangeStyle);

            m_WorkBook.setSelection(ftrRange);
            Set3DBorder(7, 1, 7, 5, Color.LightGray.ToArgb(), Color.DarkGray.ToArgb(), Color.DarkGray.ToArgb());
            AlignRight();
            m_WorkBook.setRangeStyle(m_RangeStyle);
        }

        private void Colorful1()
        {
            m_WorkBook.setSelection(1, 1, 7, 5);
            int color = Color.Red.ToArgb();
            m_RangeStyle.BottomBorder = RangeStyle.BorderThin;
            m_RangeStyle.BottomBorderColor = Color.Red.ToArgb();
            SetSolidPattern(Color.DarkGray.ToArgb(), Color.Black.ToArgb());

            m_RangeStyle.TopBorder = RangeStyle.BorderMedium;
            m_RangeStyle.LeftBorder = RangeStyle.BorderMedium;
            m_RangeStyle.RightBorder = RangeStyle.BorderMedium;
            m_RangeStyle.TopBorderColor = Color.Red.ToArgb();
            m_RangeStyle.LeftBorderColor = Color.Red.ToArgb();
            m_RangeStyle.RightBorderColor = Color.Red.ToArgb();

            AdjustFont(color, false, false, false);
            m_WorkBook.setRangeStyle(m_RangeStyle);

            m_WorkBook.setSelection(hdrRange);
            SetSolidPattern(Color.Black.ToArgb(), Color.Black.ToArgb());
            AdjustFont(color, true, true, false);
            AlignCenter();
            m_WorkBook.setRangeStyle(m_RangeStyle);

            m_WorkBook.setSelection(colRange);
            SetSolidPattern(m_WorkBook.getPaletteEntry(11), Color.Black.ToArgb());
            AdjustFont(color, true, true, false);
            m_WorkBook.setRangeStyle(m_RangeStyle);
        }

        private void Colorful2()
        {
            int color = m_WorkBook.getPaletteEntry(14);
            m_WorkBook.setSelection(hdrRange);
            m_RangeStyle.TopBorder = RangeStyle.BorderMedium;
            m_RangeStyle.BottomBorder = RangeStyle.BorderThin;
            SetSolidPattern(m_WorkBook.getPaletteEntry(9), Color.Black.ToArgb());
            AdjustFont(color, true, true, false);
            AlignRight();
            m_WorkBook.setRangeStyle(m_RangeStyle);

            m_WorkBook.setSelection(ftrRange);
            m_RangeStyle.BottomBorder = RangeStyle.BorderMedium;
            m_RangeStyle.TopBorder = RangeStyle.BorderThin;
            m_WorkBook.setRangeStyle(m_RangeStyle);

            m_WorkBook.setSelection(colRange);
            AdjustFont(Color.Black.ToArgb(), true, true, false);
            m_WorkBook.setRangeStyle(m_RangeStyle);

            m_WorkBook.setSelection(1, 1, 7, 5);
            SetHatchPattern(m_WorkBook.getPaletteEntry(16), Color.Red.ToArgb());
            m_WorkBook.setRangeStyle(m_RangeStyle);
        }

        private void Colorful3()
        {
            m_WorkBook.setSelection(1, 1, 7, 5);
            SetSolidPattern(Color.Black.ToArgb(), Color.Black.ToArgb());
            AdjustFont(Color.White.ToArgb(), false, false, false);
            m_WorkBook.setRangeStyle(m_RangeStyle);

            m_WorkBook.setSelection(hdrRange);
            AdjustFont(Color.Green.ToArgb(), true, true, false);
            AlignRight();
            m_WorkBook.setRangeStyle(m_RangeStyle);

            m_WorkBook.setSelection(colRange);
            AdjustFont(Color.Magenta.ToArgb(), true, true, false);
            m_WorkBook.setRangeStyle(m_RangeStyle);
        }

        private void List1()
        {
            m_WorkBook.setSelection(1, 1, 7, 5);
            System.String newSelection, numberFormat;
            int fcolor, bcolor;
            m_RangeStyle.TopBorder = RangeStyle.BorderThin;
            m_RangeStyle.LeftBorder = RangeStyle.BorderThin;
            m_RangeStyle.RightBorder = RangeStyle.BorderThin;
            m_RangeStyle.TopBorderColor = m_WorkBook.getPaletteEntry(16);
            m_RangeStyle.LeftBorderColor = m_WorkBook.getPaletteEntry(16);
            m_RangeStyle.RightBorderColor = m_WorkBook.getPaletteEntry(16);
            m_WorkBook.setRangeStyle(m_RangeStyle);

            m_WorkBook.setSelection(hdrRange);
            SetSolidPattern(m_WorkBook.getPaletteEntry(14), Color.Black.ToArgb());
            AdjustFont(Color.Blue.ToArgb(), true, true, false);
            AlignCenter();
            m_WorkBook.setRangeStyle(m_RangeStyle);

            m_WorkBook.setSelection(ftrRange);
            SetSolidPattern(m_WorkBook.getPaletteEntry(14), Color.Black.ToArgb());
            AdjustFont(Color.Blue.ToArgb(), true, false, false);
            numberFormat = "$ #,##0.00_);(#,##0.00)";
            m_RangeStyle.CustomFormat = numberFormat;
            AlignRight();
            m_WorkBook.setRangeStyle(m_RangeStyle);

            fcolor = m_WorkBook.getPaletteEntry(19);
            bcolor = Color.Red.ToArgb();
            for (int i = StartRow + 1; i <= EndRow - 1; i = i + 2)
            {
                newSelection = m_WorkBook.formatRCNr(i, StartCol, false) + ":" + m_WorkBook.formatRCNr(i, EndCol, false);
                m_WorkBook.setSelection(newSelection);
                SetHatchPattern(fcolor, bcolor);
                m_WorkBook.setRangeStyle(m_RangeStyle);
            }

            fcolor = Color.White.ToArgb();
            bcolor = m_WorkBook.getPaletteEntry(15);
            for (int i = StartRow + 2; i < EndRow - 1; i = i + 2)
            {
                newSelection = m_WorkBook.formatRCNr(i, StartCol, false) + ":" + m_WorkBook.formatRCNr(i, EndCol, false);
                m_WorkBook.setSelection(newSelection);
                SetHatchPattern(fcolor, bcolor);
                m_WorkBook.setRangeStyle(m_RangeStyle);
            }
        }

        private void List2()
        {
            m_WorkBook.setSelection(1, 1, 7, 5);
            System.String numberFormat, newSelection;
            int fcolor, bcolor;
            m_RangeStyle.TopBorder = RangeStyle.BorderThin;
            m_RangeStyle.LeftBorder = RangeStyle.BorderThin;
            m_RangeStyle.RightBorder = RangeStyle.BorderThin;
            m_RangeStyle.TopBorderColor = m_WorkBook.getPaletteEntry(16);
            m_RangeStyle.LeftBorderColor = m_WorkBook.getPaletteEntry(16);
            m_RangeStyle.RightBorderColor = m_WorkBook.getPaletteEntry(16);
            m_WorkBook.setRangeStyle(m_RangeStyle);

            m_WorkBook.setSelection(hdrRange);
            m_RangeStyle.TopBorder = RangeStyle.BorderThick;
            m_RangeStyle.TopBorderColor = m_WorkBook.getPaletteEntry(16);
            m_RangeStyle.BottomBorder = RangeStyle.BorderThin;
            SetSolidPattern(m_WorkBook.getPaletteEntry(15), Color.Black.ToArgb());
            AlignCenter();
            AdjustFont(Color.Red.ToArgb(), true, true, false);
            m_WorkBook.setRangeStyle(m_RangeStyle);

            m_WorkBook.setSelection(ftrRange);
            m_RangeStyle.BottomBorder = RangeStyle.BorderThick;
            m_RangeStyle.BottomBorderColor = m_WorkBook.getPaletteEntry(16);
            m_RangeStyle.TopBorder = RangeStyle.BorderThin;
            numberFormat = "$ #,##0.00_);(#,##0.00)";
            m_RangeStyle.CustomFormat = numberFormat;
            AlignRight();
            m_WorkBook.setRangeStyle(m_RangeStyle);

            fcolor = Color.Orange.ToArgb();
            bcolor = Color.White.ToArgb();
            for (int i = StartRow + 1; i < EndRow; i = i + 2)
            {
                newSelection = m_WorkBook.formatRCNr(i, StartCol, false) + ":" + m_WorkBook.formatRCNr(i, EndCol, false);

                m_WorkBook.setSelection(newSelection);
                SetHatchPattern(fcolor, bcolor);
                m_WorkBook.setRangeStyle(m_RangeStyle);
            }
        }

        private void List3()
        {
            m_WorkBook.setSelection(colRange);
            m_WorkBook.setRangeStyle(m_RangeStyle);
            m_WorkBook.setSelection(hdrRange);
            m_RangeStyle.TopBorder = RangeStyle.BorderMedium;
            m_RangeStyle.TopBorderColor = Color.DarkGray.ToArgb();
            m_RangeStyle.BottomBorder = RangeStyle.BorderMedium;
            m_RangeStyle.BottomBorderColor = Color.DarkGray.ToArgb();
            AlignCenter();
            AdjustFont(m_WorkBook.getPaletteEntry(11), true, false, false);
            m_WorkBook.setRangeStyle(m_RangeStyle);

            m_WorkBook.setSelection(ftrRange);
            m_RangeStyle = m_WorkBook.getRangeStyle(EndRow, StartCol, EndRow, EndCol);
            m_RangeStyle.TopBorder = RangeStyle.BorderMedium;
            m_RangeStyle.TopBorderColor = Color.DarkGray.ToArgb();
            m_RangeStyle.BottomBorder = RangeStyle.BorderMedium;
            m_RangeStyle.BottomBorderColor = Color.DarkGray.ToArgb();
            AlignRight();
            m_WorkBook.setRangeStyle(m_RangeStyle);
        }

        private void Set3DBorder(int row1, int col1, int row2, int col2, int outlineColor, int rightColor, int bottomColor)
        {
            m_WorkBook.setSelection(row1, col1, row2, col2);
            m_RangeStyle.TopBorder = RangeStyle.BorderMedium;
            m_RangeStyle.BottomBorder = RangeStyle.BorderMedium;
            m_RangeStyle.LeftBorder = RangeStyle.BorderMedium;
            m_RangeStyle.RightBorder = RangeStyle.BorderMedium;
            m_RangeStyle.TopBorderColor = outlineColor;
            m_RangeStyle.BottomBorderColor = outlineColor;
            m_RangeStyle.LeftBorderColor = outlineColor;
            m_RangeStyle.RightBorderColor = outlineColor;
            m_WorkBook.setRangeStyle(m_RangeStyle);

            m_WorkBook.setSelection(row1, col2, row2, col2);
            m_RangeStyle.RightBorder = RangeStyle.BorderMedium;
            m_RangeStyle.RightBorderColor = rightColor;
            m_WorkBook.setRangeStyle(m_RangeStyle);

            m_WorkBook.setSelection(row2, col1, row2, col2);
            m_RangeStyle.BottomBorder = RangeStyle.BorderMedium;
            m_RangeStyle.BottomBorderColor = bottomColor;
            m_WorkBook.setRangeStyle(m_RangeStyle);
        }

        private void SetHatchPattern(int fcolor, int bcolor)
        {
            m_RangeStyle.Pattern = (short)4;
            m_RangeStyle.PatternFG = fcolor;
            m_RangeStyle.PatternBG = bcolor;
        }

        private void AlignCenter()
        {
            m_RangeStyle.HorizontalAlignment = RangeStyle.HorizontalAlignmentCenter;
            m_RangeStyle.VerticalAlignment = RangeStyle.VerticalAlignmentBottom;
            m_RangeStyle.WordWrap = false;
        }

        private void AdjustFont(int color, bool bold, bool italic, bool underline)
        {
            m_RangeStyle.FontBold = bold;
            m_RangeStyle.FontItalic = italic;
            m_RangeStyle.FontUnderline = RangeStyle.UnderlineSingle;
            m_RangeStyle.FontColor = color;
        }

        private void AlignRight()
        {
            m_RangeStyle.HorizontalAlignment = RangeStyle.HorizontalAlignmentRight;
            m_RangeStyle.VerticalAlignment = RangeStyle.VerticalAlignmentBottom;
            m_RangeStyle.WordWrap = false;
        }

        private void SetSolidPattern(int fcolor, int bcolor)
        {
            short nPattern;
            nPattern = 1;
            m_RangeStyle.Pattern = nPattern;
            m_RangeStyle.PatternFG = fcolor;
            m_RangeStyle.PatternBG = bcolor;
        }

        string ISampleControl.CodePath
        {
            get
            {
                return "src\\Formatting\\CellFormatSample.cs";
            }
        }
    }
}

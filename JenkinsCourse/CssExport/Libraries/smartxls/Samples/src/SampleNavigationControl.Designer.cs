namespace SmartXLSExplorer
{
    partial class SampleNavigationControl
    {
        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.Windows.Forms.TreeNode treeNode1 = new System.Windows.Forms.TreeNode("ChartPreserve");
            System.Windows.Forms.TreeNode treeNode2 = new System.Windows.Forms.TreeNode("ChartEdit");
            System.Windows.Forms.TreeNode treeNode3 = new System.Windows.Forms.TreeNode("Chart");
            System.Windows.Forms.TreeNode treeNode4 = new System.Windows.Forms.TreeNode("ChartSheet");
            System.Windows.Forms.TreeNode treeNode5 = new System.Windows.Forms.TreeNode("Charts", new System.Windows.Forms.TreeNode[] {
            treeNode1,
            treeNode2,
            treeNode3,
            treeNode4});
            System.Windows.Forms.TreeNode treeNode6 = new System.Windows.Forms.TreeNode("ReadImage");
            System.Windows.Forms.TreeNode treeNode7 = new System.Windows.Forms.TreeNode("WriteImages");
            System.Windows.Forms.TreeNode treeNode8 = new System.Windows.Forms.TreeNode("Images", new System.Windows.Forms.TreeNode[] {
            treeNode6,
            treeNode7});
            System.Windows.Forms.TreeNode treeNode9 = new System.Windows.Forms.TreeNode("CellFormat");
            System.Windows.Forms.TreeNode treeNode10 = new System.Windows.Forms.TreeNode("ConditionFormat");
            System.Windows.Forms.TreeNode treeNode11 = new System.Windows.Forms.TreeNode("Formatting", new System.Windows.Forms.TreeNode[] {
            treeNode9,
            treeNode10});
            System.Windows.Forms.TreeNode treeNode12 = new System.Windows.Forms.TreeNode("Fomula");
            System.Windows.Forms.TreeNode treeNode13 = new System.Windows.Forms.TreeNode("FormulaTest");
            System.Windows.Forms.TreeNode treeNode14 = new System.Windows.Forms.TreeNode("Formulas", new System.Windows.Forms.TreeNode[] {
            treeNode12,
            treeNode13});
            System.Windows.Forms.TreeNode treeNode15 = new System.Windows.Forms.TreeNode("ImportDataTable");
            System.Windows.Forms.TreeNode treeNode16 = new System.Windows.Forms.TreeNode("ExportDataTable");
            System.Windows.Forms.TreeNode treeNode17 = new System.Windows.Forms.TreeNode("NamedRanges");
            System.Windows.Forms.TreeNode treeNode18 = new System.Windows.Forms.TreeNode("ManipulatingData", new System.Windows.Forms.TreeNode[] {
            treeNode15,
            treeNode16,
            treeNode17});
            System.Windows.Forms.TreeNode treeNode19 = new System.Windows.Forms.TreeNode("ReadWrite");
            System.Windows.Forms.TreeNode treeNode20 = new System.Windows.Forms.TreeNode("Group");
            System.Windows.Forms.TreeNode treeNode21 = new System.Windows.Forms.TreeNode("WriteHyperlinks");
            System.Windows.Forms.TreeNode treeNode22 = new System.Windows.Forms.TreeNode("ReadHyperlinks");
            System.Windows.Forms.TreeNode treeNode23 = new System.Windows.Forms.TreeNode("XLSX");
            System.Windows.Forms.TreeNode treeNode24 = new System.Windows.Forms.TreeNode("DrawingObjects");
            System.Windows.Forms.TreeNode treeNode25 = new System.Windows.Forms.TreeNode("Pivot");
            System.Windows.Forms.TreeNode treeNode26 = new System.Windows.Forms.TreeNode("Others", new System.Windows.Forms.TreeNode[] {
            treeNode19,
            treeNode20,
            treeNode21,
            treeNode22,
            treeNode23,
            treeNode24,
            treeNode25});
            System.Windows.Forms.TreeNode treeNode27 = new System.Windows.Forms.TreeNode("MultiSheetTest");
            System.Windows.Forms.TreeNode treeNode28 = new System.Windows.Forms.TreeNode("RangeCopy");
            System.Windows.Forms.TreeNode treeNode29 = new System.Windows.Forms.TreeNode("Freezepane");
            System.Windows.Forms.TreeNode treeNode30 = new System.Windows.Forms.TreeNode("HideUnhideRowsandColumns");
            System.Windows.Forms.TreeNode treeNode31 = new System.Windows.Forms.TreeNode("PageBreak");
            System.Windows.Forms.TreeNode treeNode32 = new System.Windows.Forms.TreeNode("RowHeightandColumnWidth");
            System.Windows.Forms.TreeNode treeNode33 = new System.Windows.Forms.TreeNode("WorksheetManipulation", new System.Windows.Forms.TreeNode[] {
            treeNode27,
            treeNode28,
            treeNode29,
            treeNode30,
            treeNode31,
            treeNode32});
            System.Windows.Forms.TreeNode treeNode34 = new System.Windows.Forms.TreeNode("RichTextWrite");
            System.Windows.Forms.TreeNode treeNode35 = new System.Windows.Forms.TreeNode("RichTextRead");
            System.Windows.Forms.TreeNode treeNode36 = new System.Windows.Forms.TreeNode("RichString", new System.Windows.Forms.TreeNode[] {
            treeNode34,
            treeNode35});
            System.Windows.Forms.TreeNode treeNode37 = new System.Windows.Forms.TreeNode("CommentWrite");
            System.Windows.Forms.TreeNode treeNode38 = new System.Windows.Forms.TreeNode("CommentRead");
            System.Windows.Forms.TreeNode treeNode39 = new System.Windows.Forms.TreeNode("Comment", new System.Windows.Forms.TreeNode[] {
            treeNode37,
            treeNode38});
            System.Windows.Forms.TreeNode treeNode40 = new System.Windows.Forms.TreeNode("DataValidationRead");
            System.Windows.Forms.TreeNode treeNode41 = new System.Windows.Forms.TreeNode("DataValidationWrite");
            System.Windows.Forms.TreeNode treeNode42 = new System.Windows.Forms.TreeNode("DataValidation", new System.Windows.Forms.TreeNode[] {
            treeNode40,
            treeNode41});
            System.Windows.Forms.TreeNode treeNode43 = new System.Windows.Forms.TreeNode("SmartXLSExplorer", new System.Windows.Forms.TreeNode[] {
            treeNode5,
            treeNode8,
            treeNode11,
            treeNode14,
            treeNode18,
            treeNode26,
            treeNode33,
            treeNode36,
            treeNode39,
            treeNode42});
            this.treeView1 = new System.Windows.Forms.TreeView();
            this.SuspendLayout();
            // 
            // treeView1
            // 
            this.treeView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.treeView1.HideSelection = false;
            this.treeView1.HotTracking = true;
            this.treeView1.Location = new System.Drawing.Point(0, 0);
            this.treeView1.Margin = new System.Windows.Forms.Padding(2);
            this.treeView1.Name = "treeView1";
            treeNode1.Name = "ChartPreserveSample";
            treeNode1.Text = "ChartPreserve";
            treeNode2.Name = "ChartEditSample";
            treeNode2.Text = "ChartEdit";
            treeNode3.Name = "ChartSample";
            treeNode3.Text = "Chart";
            treeNode4.Name = "ChartSheetSample";
            treeNode4.Text = "ChartSheet";
            treeNode5.Name = "ChartsCategory";
            treeNode5.Text = "Charts";
            treeNode6.Name = "ReadImageSample";
            treeNode6.Text = "ReadImage";
            treeNode7.Name = "WriteImagesSample";
            treeNode7.Text = "WriteImages";
            treeNode8.Name = "ImagesCategory";
            treeNode8.Text = "Images";
            treeNode9.Name = "CellFormatSample";
            treeNode9.Text = "CellFormat";
            treeNode10.Name = "ConditionFormatSample";
            treeNode10.Text = "ConditionFormat";
            treeNode11.Name = "FormattingCategory";
            treeNode11.Text = "Formatting";
            treeNode12.Name = "FomulaSample";
            treeNode12.Text = "Fomula";
            treeNode13.Name = "FormulaTestSample";
            treeNode13.Text = "FormulaTest";
            treeNode14.Name = "FormulasCategory";
            treeNode14.Text = "Formulas";
            treeNode15.Name = "ImportDataTableSample";
            treeNode15.Text = "ImportDataTable";
            treeNode16.Name = "ExportDataTableSample";
            treeNode16.Text = "ExportDataTable";
            treeNode17.Name = "NamedRangesSample";
            treeNode17.Text = "NamedRanges";
            treeNode18.Name = "ManipulatingDataCategory";
            treeNode18.Text = "ManipulatingData";
            treeNode19.Name = "ReadWriteSample";
            treeNode19.Text = "ReadWrite";
            treeNode20.Name = "GroupSample";
            treeNode20.Text = "Group";
            treeNode21.Name = "WriteHyperlinksSample";
            treeNode21.Text = "WriteHyperlinks";
            treeNode22.Name = "ReadHyperlinksSample";
            treeNode22.Text = "ReadHyperlinks";
            treeNode23.Name = "XLSXSample";
            treeNode23.Text = "XLSX";
            treeNode24.Name = "DrawingObjectsSample";
            treeNode24.Text = "DrawingObjects";
            treeNode25.Name = "PivotSample";
            treeNode25.Text = "Pivot";
            treeNode26.Name = "OthersCategory";
            treeNode26.Text = "Others";
            treeNode27.Name = "MultiSheetTestSample";
            treeNode27.Text = "MultiSheetTest";
            treeNode28.Name = "RangeCopySample";
            treeNode28.Text = "RangeCopy";
            treeNode29.Name = "FreezepaneSample";
            treeNode29.Text = "Freezepane";
            treeNode30.Name = "HideUnhideRowsandColumnsSample";
            treeNode30.Text = "HideUnhideRowsandColumns";
            treeNode31.Name = "PageBreakSample";
            treeNode31.Text = "PageBreak";
            treeNode32.Name = "RowHeightandColumnWidthSample";
            treeNode32.Text = "RowHeightandColumnWidth";
            treeNode33.Name = "ManipulationCategory";
            treeNode33.Text = "WorksheetManipulation";
            treeNode34.Name = "RichTextWriteSample";
            treeNode34.Text = "RichTextWrite";
            treeNode35.Name = "RichTextReadSample";
            treeNode35.Text = "RichTextRead";
            treeNode36.Name = "RichStringCategory";
            treeNode36.Text = "RichString";
            treeNode37.Name = "CommentWriteSample";
            treeNode37.Text = "CommentWrite";
            treeNode38.Name = "CommentReadSample";
            treeNode38.Text = "CommentRead";
            treeNode39.Name = "CommentCategory";
            treeNode39.Text = "Comment";
            treeNode40.Name = "DataValidationReadSample";
            treeNode40.Text = "DataValidationRead";
            treeNode41.Name = "DataValidationWriteSample";
            treeNode41.Text = "DataValidationWrite";
            treeNode42.Name = "DataValidationCategory";
            treeNode42.Text = "DataValidation";
            treeNode43.Name = "SmartXLSCategory";
            treeNode43.Text = "SmartXLSExplorer";
            this.treeView1.Nodes.AddRange(new System.Windows.Forms.TreeNode[] {
            treeNode43});
            this.treeView1.Size = new System.Drawing.Size(197, 236);
            this.treeView1.TabIndex = 0;
            this.treeView1.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.treeView1_AfterSelect);
            // 
            // SampleNavigationControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.treeView1);
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "SampleNavigationControl";
            this.Size = new System.Drawing.Size(197, 236);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TreeView treeView1;

    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using NUnit.Framework;
using CssExport;
using Path = System.IO.Path;



namespace TestCssExport
{
    [TestFixture]
    public class TestCssExport
    {
        [TestCase]
        public void GetContentOfCsvFileAsArray()
        {
            string dir = FileUtils.GetBrowserDownloadsDirectoryPath();
            string fileNameWithFullPath = Path.Combine(dir, "Exported.csv");
            string[,] fileContent = FileUtils.GetContentOfCsvFileAsArray(fileNameWithFullPath);

            //string[] rows = FileUtils.GetFileMultipleEntry(fileNameWithFullPath); 

            //KeyValuePair<int, int> numberOfRowsAndColumns = FileUtils.GetNumberOfRowsAndColumnsFromCsvFile(fileNameWithFullPath);
            //string[,] fileContent = new string[numberOfRowsAndColumns.Key, numberOfRowsAndColumns.Value];
            //foreach(var row in rows.Select((v,i) => new {Value = v, Index = i}))
            //{
            //    string[] columns = row.Value.ToString().Split(',');            
            //    foreach (var column in columns.Select((v, i) => new { Value = v, Index = i }))
            //    {
            //        fileContent[row.Index,column.Index] = columns[column.Index];
            //    }
            //}

            Assert.AreEqual("vv", "aa");
        }
    }
}

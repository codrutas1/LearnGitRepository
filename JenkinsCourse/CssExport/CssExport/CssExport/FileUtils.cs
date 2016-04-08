using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

using Excel;

using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;

using Microsoft.VisualBasic.FileIO;

using Path = System.IO.Path;

//namespace POMAutomationTests.Infrastructure.Utils
namespace CssExport
{
    public static class FileUtils
    {
        private const string c_dateRegularExpression = @"[0-9]{1,2}\/[0-9]{1,2}\/[0-9]{4}";

        private const string c_noDateStringFoundMessage = "No date string has been found in the selected text file.";

        /// <summary>Deletes all the files and directories from a directory. If the directory does not exist, it does nothing</summary>
        /// <param name="directoryPath">The path of the directory</param>
        public static void DeleteAllFilesAndDirectoriesFromADirectory(string directoryPath)
        {
            if (Directory.Exists(directoryPath))
            {
                Array.ForEach(Directory.GetFiles(directoryPath), File.Delete);
                foreach (string directory in Directory.GetDirectories(directoryPath))
                {
                    Directory.Delete(directory, true);
                }
            }
        }

        /// <summary>Returns the full file path to the Downloads directory</summary>
        /// <returns>The Downloads directory path</returns>
        public static string GetBrowserDownloadsDirectoryPath()
        {
            return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "BrowserDownloads");
        }

        /// <summary>Getting the first date string object in short format, found in the given text file, like "8/31/2014"</summary>
        /// <param name="fileName">The given file name to extract the date from</param>
        /// <param name="index">Index for the occurrence</param>
        /// <returns>The retrieved date string</returns>
        public static string GetDateStringFromResourceFileInShortFormat(string fileName, int index = 0)
        {
            string fileContent = FileUtils.GetContentOfResourceFile(fileName);
            string foundDate;

            MatchCollection match = Regex.Matches(fileContent, FileUtils.c_dateRegularExpression);

            if (match[index].Success)
            {
                foundDate = match[index].Value;
            }
            else
            {
                foundDate = FileUtils.c_noDateStringFoundMessage;
            }

            return foundDate;
        }

        /// <summary>Searching for a given string parameter in a given file and replacing it in another file</summary>
        /// <param name="originalFileName">The file name containing the string to be replaced</param>
        /// <param name="temporaryFileName">The file name to be generated with the modified content</param>
        /// <param name="stringToReplace">The string to be replaced</param>
        /// <param name="stringToReplaceWith">The string to replace the original string</param>
        public static void ReplaceStringInResourceFile(string originalFileName, string temporaryFileName, string stringToReplace, string stringToReplaceWith)
        {
            string originalText = FileUtils.GetContentOfResourceFile(originalFileName);
            string modifiedText = originalText.Replace(stringToReplace, stringToReplaceWith);

            File.WriteAllText(FileUtils.ResolveResourceFilePath(temporaryFileName), modifiedText);
        }

        /// <summary>Deletes a given file from the BaseDirectory</summary>
        /// <param name="fileName">The file name from the BaseDirecotry to be delete</param>
        /// <returns>The delete operation status message</returns>
        public static void DeleteResourceFile(string fileName)
        {
            string filePath = FileUtils.ResolveResourceFilePath(fileName);

            File.Delete(filePath);
        }

        /// <summary>Gets the number of lines of a text file</summary>
        /// <param name="fileName">The name of the text file</param>
        /// <returns></returns>
        public static int GetNumberOfLinesFromTextResourceFile(string fileName)
        {
            string filePath = FileUtils.ResolveResourceFilePath(fileName);
            return File.ReadLines(filePath).Count();
        }


        /// <summary>Returns the number of rows and columns of a csv file</summary>
        /// <param name="fileName">The name of the text file</param>
        /// <returns>A pair containing first the number of rows and second the number of columns</returns>
        public static KeyValuePair<int, int> GetNumberOfRowsAndColumnsFromCsvFile(string fileName)
        {
            int numberOfRows = File.ReadLines(fileName).Count();
            string[] rows = FileUtils.GetFileMultipleEntry(fileName);
            int numberOfColumns = rows[0].Split(',').Length;
            return new KeyValuePair<int, int>(numberOfRows, numberOfColumns);
        }


         /// <summary>Returns the content of a csv file as a two-dimension array</summary>
        /// <param name="fileName">The name of the text file</param>
        /// <returns>The two-dimension array</returns>
        public static string[,] GetContentOfCsvFileAsArray(string fileName)
        {

            //string dir = FileUtils.GetBrowserDownloadsDirectoryPath();
            //string fileNameWithFullPath = Path.Combine(dir, "Exported.csv");
            string[] rows = FileUtils.GetFileMultipleEntry(fileName); 
            KeyValuePair<int, int> numberOfRowsAndColumns = FileUtils.GetNumberOfRowsAndColumnsFromCsvFile(fileName);
            string[,] fileContent = new string[numberOfRowsAndColumns.Key, numberOfRowsAndColumns.Value];

            foreach(var row in rows.Select((v,i) => new {Value = v, Index = i}))
            {
                string[] columns = row.Value.ToString().Split(',');            
                foreach (var column in columns.Select((v, i) => new { Value = v, Index = i }))
                {
                    fileContent[row.Index,column.Index] = columns[column.Index];
                }              
            }
            return fileContent;
        }


        ///// <summary>Retrurns a new file name based on the one received as parameter. THe new file name has a unique timestamp appended.</summary>
        ///// <param name="fileName">The file name from which to </param>
        ///// <returns></returns>
        //public static string GetNewFileNameFrom(string fileName)
        //{
        //    string uniqueId = "_" + Utils.GetUniqueTimestamp();
        //    int extensionIndex = fileName.LastIndexOf('.');
        //    string newFileName = extensionIndex < 0 ? fileName + Utils.GetUniqueTimestamp() : fileName.Substring(0, extensionIndex) + uniqueId + "." + fileName.Substring(extensionIndex + 1);
        //    return newFileName;
        //}

        ///// <summary>Copies the given file under a new name</summary>
        ///// <param name="fileName"></param>
        ///// <returns>The new created file</returns>
        //public static string CreateNewFileFrom(string fileName)
        //{
        //    string newFileName = FileUtils.GetNewFileNameFrom(fileName);
        //    FileUtils.SaveFileTo(fileName, newFileName);
        //    return newFileName;
        //}

        /// <summary>Returns the name of the last added file in a directory</summary>
        /// <param name="directoryPath">The path of the directory to search of</param>
        /// <returns>The name of the last added file</returns>
        public static string GetLastFileAddedToADirectory(string directoryPath)
        {
            return new DirectoryInfo(directoryPath).GetFiles().OrderByDescending(f => f.LastWriteTime).First().Name;
        }

        /// <summary>Saves the specified content to a file. If the file already exists it is overwritten</summary>
        /// <param name="fileName">The name of the file</param>
        /// <param name="content">The content to be saved/overwritten/</param>
        public static void WriteContentOfResourceFile(string fileName, string content)
        {
            string path = FileUtils.ResolveResourceFilePath(fileName);
            File.WriteAllText(path, content);
        }

        /// <summary>Delets all the files from a directory. If the directory does not exist, it does nothing</summary>
        /// <param name="directoryPath">The path of the directory</param>
        public static void DeleteAllFilesFromADirectory(string directoryPath)
        {
            if (Directory.Exists(directoryPath))
            {
                Array.ForEach(Directory.GetFiles(directoryPath), File.Delete);
            }
        }

        ///// <summary>Extracts all the files in the specified zip archive to a directory on the file system.</summary>
        ///// <param name="sourceArchiveFileName">The path to the archive that is to be extracted. </param>
        ///// <param name="destinationDirectoryName">
        ///// The path to the directory in which to place the extracted files, specified as a relative or absolute path. A
        ///// relative path is interpreted as relative to the current working directory.
        ///// </param>
        //public static void ExtractZipFile(string sourceArchiveFileName, string destinationDirectoryName)
        //{
        //    ZipFile.ExtractToDirectory(sourceArchiveFileName, destinationDirectoryName);
        //}

        /// <summary>Returns the full file path to the specified file that is used in tests.</summary>
        /// <param name="fileName">The name of the file.</param>
        /// <returns>THe full path to the file, including the file.</returns>
        public static string ResolveResourceFilePath(string fileName)
        {
            return Path.Combine(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources"), fileName);
        }

        /// <summary>Method to open file and return contents in an array of strings by line</summary>
        /// <param name="filePath"></param>
        /// <returns>Returns array of strings containing file contents</returns>
        public static string[] GetFileMultipleEntry(string filePath)
        {
            string[] contents = File.ReadAllLines(filePath);
            return contents;
        }

        /// <summary>Method to parse a single line file and return the contents as an array</summary>
        /// <param name="inputFile"></param>
        /// <returns>Returns array of strings containing file contents</returns>
        public static string[] ParseSingleEntryResourceFile(string inputFile)
        {
            string[] parsedFile = FileUtils.ParseStringtoArray(FileUtils.GetFileSingleEntry(FileUtils.ResolveResourceFilePath(inputFile)));
            return parsedFile;
        }


        /// <summary>Method to parse a multi-line file and return the contents as an array</summary>
        /// <param name="inputFile"></param>
        /// <returns>Returns array of strings containing array of strings for each line of file</returns>
        public static string[][] ParseMultipleEntryResourceFile(string inputFile)
        {
            string[][] parsedFile = FileUtils.ParseStringtoArray(FileUtils.GetFileMultipleEntry(FileUtils.ResolveResourceFilePath(inputFile)));
            return parsedFile;
        }

        /// <summary>Method to parse a tab delimited file and return the contents as a List</summary>
        /// <param name="inputFile">The given file name to retrieve the content from</param>
        /// <returns>Returns a list of strings for each column of file</returns>
        public static List<string> ParseTabDelimitedResourceFileToList(string inputFile)
        {
            var parsedList = new List<string>();
            string[] lines = File.ReadAllLines(FileUtils.ResolveResourceFilePath(inputFile));
            foreach (string line in lines)
            {
                string[] columns = line.Split('\t');
                columns.ToList().ForEach(q => parsedList.Add(q));
            }
            return parsedList;
        }

        /// <summary>Returns a string representing the content of a given file</summary>
        /// <param name="fileName">The given file name to retrieve the content from</param>
        /// <returns>A string representing the content of the given file</returns>
        public static string GetContentOfResourceFile(string fileName)
        {
            string filePath = FileUtils.ResolveResourceFilePath(fileName);
            return File.ReadAllText(filePath);
        }

        /// <summary>Parses an excel file to get a list of strings containing a specific column from the file</summary>
        /// <param name="fileName">The given file name to retrieve the content from</param>
        /// <param name="excelSheetNumber">The number of the sheet in the file, with index 1</param>
        /// <param name="startingRowNumber">The number of the row where to start getting the data from, with index 1</param>
        /// <param name="columnNumber">The number of the column, with index 1</param>
        /// <returns>A list of string containing the fields content</returns>
        public static List<string> GetColumnContentOfExcelFile(string fileName, int excelSheetNumber, int startingRowNumber, int columnNumber)
        {
            using (FileStream stream = File.Open(fileName, FileMode.Open, FileAccess.Read))
            {
                using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream))
                {
                    DataSet result = excelReader.AsDataSet();

                    DataRowCollection totalRows = result.Tables[excelSheetNumber - 1].Rows;
                    var listOfStrings = new List<string>();
                    for (int i = startingRowNumber - 1; i < totalRows.Count; i++)
                    {
                        string fieldContent = result.Tables[excelSheetNumber - 1].Rows[i].ItemArray[columnNumber - 1].ToString();
                        listOfStrings.Add(fieldContent);
                    }
                    return listOfStrings;
                }
            }
        }

        /// <summary>Gets the first line of an excel file; needed when a list of table headers from an excel file is needed;</summary>
        /// <param name="fileName"></param>
        /// <param name="excelSheetNumber"></param>
        /// <returns></returns>
        public static List<string> GetFirstLineOfExcelFile(string fileName, int excelSheetNumber)
        {
            using (FileStream stream = File.Open(fileName, FileMode.Open, FileAccess.Read))
            {
                using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream))
                {
                    DataSet result = excelReader.AsDataSet();
                    return result.Tables[excelSheetNumber - 1].Rows[0].ItemArray.Select(t => t.ToString()).ToList();
                }
            }
        }

        /// <summary>Opens and parses a CSV file and returns a list with the content</summary>
        /// <param name="filePath">The path of the CSV file</param>
        /// <returns>list with content of the CSV file</returns>
        public static List<string> GetContentOfCsvFile(string filePath)
        {
            var parser = new TextFieldParser(filePath);
            parser.TextFieldType = FieldType.Delimited;
            parser.SetDelimiters(",", "\t");
            var listOfStrings = new List<string>();
            while (!parser.EndOfData)
            {
                //Process row
                string[] fields = parser.ReadFields();
                foreach (string field in fields)
                {
                    listOfStrings.Add(field);
                }
            }
            parser.Close();
            return listOfStrings;
        }

        /// <summary>Opens and parses a CSV file and returns a dictonary with the content</summary>
        /// <param name="filePath">The path of the CSV file</param>
        /// <returns>Dictionary with first item in each row as "key" or first item & rest of items in array</returns>
        public static Dictionary<string, string[]> ParseMultiLineCsvWithKey(string filePath)
        {
            var table = new Dictionary<string, string[]>();
            var parser = new TextFieldParser(filePath);
            parser.TextFieldType = FieldType.Delimited;
            parser.SetDelimiters(",");
            while (!parser.EndOfData)
            {
                //Process row
                string[] fields = parser.ReadFields();
                table[fields[0]] = fields.Skip(1).ToArray();
            }
            parser.Close();
            return table;
        }

        /// <summary>Runs the sql with the specified tokens/parameters. </summary>
        /// <param name="sqlStatement">The sql string</param>
        /// <param name="connectionString">The connection string</param>
        /// <param name="sqlParameters">The sql parameters to be passed into the DB script</param>
        /// <returns>Dictionary with first item in each row as "key" or first item & rest of items in array</returns>
        public static DataTable ReturnDbQueryResultsWithKey(Dictionary<string, string> sqlParameters, string connectionString, string sqlStatement)
        {
            return FileUtils.ExecuteSql(connectionString, sqlStatement, sqlParameters.Select(p => new SqlParameter(p.Key, p.Value)).ToArray());
        }

        /// <summary>Extracts the text from a PDF file line by line from all the pages</summary>
        /// <param name="path">The path of the PDF file</param>
        /// <returns>The list of the PDF lines</returns>
        public static List<string> ExtractTextFromPdf(string path)
        {
            using (var reader = new PdfReader(path))
            {
                var text = new StringBuilder();
                for (int i = 1; i <= reader.NumberOfPages; i++)
                {
                    text.Append(PdfTextExtractor.GetTextFromPage(reader, i));
                }
                return text.ToString().Split(new[] { "\r\n", "\n" }, StringSplitOptions.None).ToList();
            }
        }

        /// <summary>Method to convert string[] to string[][]</summary>
        /// <param name="inputArray"></param>
        /// <returns>Returns array of string[] separating string contents into own array segments</returns>
        private static string[][] ParseStringtoArray(string[] inputArray)
        {
            var contents = new string[inputArray.Length][];
            for (int i = 0; i < inputArray.Length; i++)
            {
                contents[i] = inputArray[i].Split(null);
            }
            return contents;
        }

        /// <summary>Method to convert string to string[]</summary>
        /// <param name="inputString"></param>
        /// <returns>Returns array of strings separating string contents into own array segments</returns>
        private static string[] ParseStringtoArray(string inputString)
        {
            string[] contents = inputString.Split(null);
            return contents;
        }

        /// <summary>Saves a file under a new name</summary>
        /// <param name="fileName">File co copy the content from</param>
        /// <param name="newFileName">New file to be created having the same content as the first file</param>
        private static void SaveFileTo(string fileName, string newFileName)
        {
            FileUtils.WriteContentOfResourceFile(newFileName, FileUtils.GetContentOfResourceFile(fileName));
        }

        /// <summary>Method to open file and return contents in single string</summary>
        /// <param name="filePath"></param>
        /// <returns>Return string containing file contents</returns>
        private static string GetFileSingleEntry(string filePath)
        {
            string contents = File.ReadAllText(filePath);
            return contents;
        }

        private static DataTable ExecuteSql(string connectionString, string sqlStatement, SqlParameter[] sqlParameters)
        {
            using (var sqlCon = new SqlConnection(connectionString))
            {
                var sqlCmd = new SqlCommand(sqlStatement, sqlCon)
                {
                    CommandTimeout = 600 // seconds
                };
                sqlCmd.Parameters.AddRange(sqlParameters);
                var dataAdapter = new SqlDataAdapter(sqlCmd);
                var dtResult = new DataTable();
                dataAdapter.Fill(dtResult);
                return dtResult;
            }
        }
    }
}



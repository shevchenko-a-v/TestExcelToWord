using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Text.RegularExpressions;

using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace TestExcelToWord
{
    public class ExcelToWordConverter
    {
        public void TransferExcelToWord(string sourceFilePath, string outputFilePath)
        {
            try
            {
                _log.Clear();// reset log
                WriteLog($@"Started transfer from [{sourceFilePath}] to [{outputFilePath}]");
                if (!File.Exists(sourceFilePath))
                    throw new FileNotFoundException("Source Excel file does not exist.", sourceFilePath);
                if (File.Exists(outputFilePath))
                {
                    WriteLog("Removing destination file");
                    File.Delete(outputFilePath);
                    WriteLog("Destination file is successfully removed.");
                }

                string letters, digits;
                ReadFromExcel(sourceFilePath, out letters, out digits);
                SaveToWordFile(outputFilePath, letters, digits);
            }
            catch (Exception e)
            {
                WriteLog(e.Message);
                throw;
            }
        }

        private void ReadFromExcel(string filePath, out string letters, out string digits)
        {
            Excel.Workbook wb = null;
            try
            {
                WriteLog("Started reading from source file.");
                var excel = new Excel.Application();
                if (excel == null)
                    throw new InvalidOperationException("Excel could not be started. Check that you have Microsoft Office installed.");
                WriteLog("Excel is started.");

                wb = excel.Workbooks.Open(filePath, 0, true); // open workbook
                if (wb == null)
                    throw new FileLoadException("Could not open workbook.", filePath);
                WriteLog("Workbook is successfully opened.");
                var sheets = wb.Worksheets;
                var ws = (Excel.Worksheet)sheets.get_Item(1);
                WriteLog("First worksheet is obtained.");

                var firstColumn = ws.UsedRange.Columns[1];
                var myvalues = (Array)firstColumn.Cells.Value;
                WriteLog("First column values are obtained.");
                string[] strArray = myvalues.OfType<object>().Select(o => o.ToString()).ToArray();

                Regex onlyLettersRegex = new Regex(@"^\p{L}+$");
                Regex onlyDigitsRegex = new Regex(@"^\d+$");
                letters = strArray.Where(x => onlyLettersRegex.IsMatch(x))
                                            .Aggregate((all, cur) => all + " " + cur);
                WriteLog("Letters cells have been formed into single string.");
                digits = strArray.Where(x => onlyDigitsRegex.IsMatch(x))
                                           .Select(x => x.Truncate(MaxDigitLength))
                                           .Aggregate((all, cur) => all + "-" + cur); ;
                WriteLog("Digits cells have been formed into single string.");
            }
            catch
            {
                WriteLog("Error occured during reading from source file.");
                throw;
            }
            finally
            {
                if (wb != null)
                {
                    wb.Close(false); // close workbook
                }
            }
            WriteLog("Reading from source file completed successfully.");
        }

        private void SaveToWordFile(string filePath, string letters, string digits)
        {
            Word.Document doc = null;
            try
            {
                WriteLog("Started writing to destination file.");
                object missing = System.Reflection.Missing.Value;
                var word = new Word.Application();
                if (word == null)
                    throw new InvalidOperationException("Word could not be started. Check that you have Microsoft Office installed.");
                WriteLog("Word is started.");
                doc = word.Documents.Add(ref missing, ref missing, ref missing, ref missing);
                WriteLog("New document created.");
                doc.Content.Font.Size = 12;
                doc.Content.Text = letters + System.Environment.NewLine + digits;
                WriteLog("Text added.");
                
                doc.SaveAs2(filePath);
                WriteLog("Word file has been saved to disk.");
            }
            catch
            {
                WriteLog("Error occured during writing to destination file.");
                throw;
            }
            finally
            {
                doc.Close();
            }
            WriteLog("Writing to destination file completed successfully.");
        }

        private void WriteLog(string message)
        {
            string timeStamp = DateTime.Now.ToString("HH:mm:ss.fff");
            _log.AppendLine(string.Format($@"{timeStamp}    {message}"));
        }

        public string Log { get { return _log.ToString(); } }
        
        private StringBuilder _log = new StringBuilder();
        private const int MaxDigitLength = 4;
    }
}

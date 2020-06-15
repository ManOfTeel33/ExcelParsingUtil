using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelParsingUtil
{
    public static class ExcelUtils
    {
        // Read Excel data to generic list - overloaded version 1
        public static IList<T> GetDataToList<T>(byte[] excelArr, Func<IList<string>, IList<string>, T> addProductData)
        {
            return GetDataToList<T>(excelArr, "", addProductData);
        }

        // Read Excel data to generic list - overloaded version 2
        public static IList<T> GetDataToList<T>(byte[] excelArr, string sheetName, Func<IList<string>, IList<string>, T> addData)
        {
            var resultList = new List<T>();
            Stream stream = new MemoryStream(excelArr);

            // Open the spreadsheet document for read-only access
            using (var document = SpreadsheetDocument.Open(stream, false))
            {
                var wbPart = document.WorkbookPart;
                var sheet = wbPart.Workbook.Descendants<Sheet>().FirstOrDefault(s =>
                    string.IsNullOrWhiteSpace(sheetName) || s.Name == sheetName);
                var wsPart = sheet != null ? (WorksheetPart)wbPart.GetPartById(sheet.Id) : null;
                if (wsPart == null)
                {                  
                    throw new Exception("No worksheet.");
                }

                // List to hold custom column names for mapping data to columns (index-free)
                var columnNames = new List<string>();

                // List to hold column address letters for handling empty cells
                var columnLetters = new List<string>();

                // Iterate cells of custom header row
                foreach (Cell cell in wsPart.Worksheet.Descendants<Row>().ElementAt(0))
                {
                    // Get custom column names
                    // Remove spaces, symbols (except underscore), and make lower cases and for all values in columnNames list                 
                    columnNames.Add(Regex.Replace(GetCellValue(document, cell), @"[^A-Za-z0-9_]", "").ToLower());

                    // Get built-in column names by extracting letters from cell references
                    columnLetters.Add(GetColumnAddress(cell.CellReference));
                }

                foreach (var row in GetUsedRows(document, wsPart))
                {
                    // Used for sheet row data to be added
                    var rowData = new List<string>();

                    rowData.AddRange(GetCellsForRow(row, columnLetters).Select(cell => GetCellValue(document, cell)));

                    // Add to the list
                    resultList.Add(addData(rowData, columnNames));
                }
            }
            return resultList;
        }

        private static string GetCellValue(SpreadsheetDocument document, Cell cell)
        {
            if (cell == null)
            {
                return null;
            }
            var value = cell.InnerText;

            // Process values particularly for those data types
            if (cell.DataType != null)
            {
                switch (cell.DataType.Value)
                {
                    // Obtain values from shared string table
                    case CellValues.SharedString:
                        var sstPart = document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                        if (sstPart != null)
                        {
                            value = sstPart.SharedStringTable.ChildElements[int.Parse(value)].InnerText;
                        }
                        break;
                }
            }
            return value;
        }

        private static IEnumerable<Row> GetUsedRows(SpreadsheetDocument document, WorksheetPart wsPart)
        {
            // Iterate all rows except the first one which should be column headers
            return from row in wsPart.Worksheet.Descendants<Row>().Skip(1) let hasValue = row.Descendants<Cell>().Any(cell => !string.IsNullOrEmpty(GetCellValue(document, cell))) where hasValue select row;
        }

        private static IEnumerable<Cell> GetCellsForRow(Row row, List<string> columnLetters)
        {
            var workIdx = 0;        
            foreach (var cell in row.Descendants<Cell>())
            {
                // Get letter of the cell address
                var cellLetter = GetColumnAddress(cell.CellReference);

                // Get column index of the matched cell
                var currentActualIdx = columnLetters.IndexOf(cellLetter);
                var emptyCell = new Cell { DataType = null, CellValue = new CellValue(string.Empty) };

                // Add empty cell if work index smaller than actual index
                for (; workIdx < currentActualIdx; workIdx++)
                {
                    yield return emptyCell;
                }

                yield return cell;
                workIdx++;

                if (cell == row.LastChild)
                {
                    //Append empty cells to enumerable. 
                    for (; workIdx < columnLetters.Count(); workIdx++)
                    {
                        yield return emptyCell;
                    }
                }          
            }                
        }

        private static string GetColumnAddress(string cellReference)
        {
            // Create a regular expression to get column address letters
            var regex = new Regex("[A-Za-z]+");
            var match = regex.Match(cellReference);
            return match.Value;
        }
    }
}
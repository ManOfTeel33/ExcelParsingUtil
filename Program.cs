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
        public static IList<T> GetDataToList<T>(SpreadsheetDocument document, Func<IList<string>, IList<string>, int, T> addData, bool hasHeaders = true)
        {
            return GetDataToList<T>(document, "", addData, hasHeaders);
        }

        public static IList<T> GetDataToList<T>(SpreadsheetDocument document, string sheetName, Func<IList<string>, IList<string>, int, T> addData, bool hasHeaders = true)
        {
            var resultList = new List<T>();

            var wsPart = GetWorksheetPartFromDocument(document, sheetName);
            if (wsPart == null)
            {                  
                throw new Exception("No worksheet.");
            }

            var firstRow = GetRow(wsPart, 0);
            var columnLetters = GetCellLetters(firstRow);
            var columnNames = columnLetters;
            if (hasHeaders)
            {
                columnNames = GetCellValueAsColumnName(document, firstRow);
            }

            foreach (var row in GetUsedRows(document, wsPart))
            {
                var rowData = new List<string>();
                rowData.AddRange(GetCellsForRow(row, columnLetters).Select(cell => GetCellValue(document, cell)));
                resultList.Add(addData(rowData, columnNames, int.Parse(row.RowIndex)));
            }
            return resultList;
        }

        public static SpreadsheetDocument OpenDocumentFromByteArray(byte[] excelArr, bool editable = false)
        {
            Stream stream = new MemoryStream(excelArr);
            return SpreadsheetDocument.Open(stream, editable);
        }

        public static WorksheetPart GetWorksheetPartFromDocument(SpreadsheetDocument document, string sheetName)
        {
            var wbPart = document.WorkbookPart;
            var sheet = wbPart.Workbook.Descendants<Sheet>().FirstOrDefault(s =>
                string.IsNullOrWhiteSpace(sheetName) || s.Name == sheetName);
           return sheet != null ? (WorksheetPart)wbPart.GetPartById(sheet.Id) : null;
        }

        public static IList<string> GetCellValueAsColumnName(SpreadsheetDocument document, Row row)
        {
            return (
                from Cell cell
                in row
                select Regex.Replace(
                GetCellValue(document, cell),
                @"[^A-Za-z0-9_]",
                ""
                ).ToLower()
            )
                .TakeWhile(
                    columnName => columnName.Length >= 1
                )
                .ToList();
        }

        public static IList<string> GetCellLetters(Row row)
        {
            return (from Cell cell in row select GetColumnAddress(cell.CellReference)).ToList();
        }

        public static string GetCellValue(SpreadsheetDocument document, Cell cell)
        {
            if (cell == null)
            {
                return null;
            }
            var value = cell?.CellValue?.InnerText ?? cell.InnerText;

            if (cell.DataType == null)
            {
                return value;
            }
            switch (cell.DataType.Value)
            {
                case CellValues.SharedString:
                    var sstPart = document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                    if (sstPart != null)
                    {
                        value = sstPart.SharedStringTable.ChildElements[int.Parse(value)].InnerText;
                    }
                    break;
            }
            return value;
        }

        public static bool HasAnyRows(WorksheetPart wsPart)
        {
            return wsPart.Worksheet.Descendants<Row>().Any();
        }

        public static Row GetRow(WorksheetPart wsPart, int rowNumber)
        {
            return wsPart.Worksheet.Descendants<Row>().ElementAt(rowNumber);
        }

        private static IEnumerable<Row> GetUsedRows(SpreadsheetDocument document, WorksheetPart wsPart)
        {
            return 
                from row 
                in wsPart.Worksheet.Descendants<Row>().Skip(1) 
                let hasValue = row.Descendants<Cell>().Any(
                    cell => !string.IsNullOrEmpty(GetCellValue(document, cell))
                    ) 
                where hasValue 
                select row;
        }

        public static IEnumerable<Cell> GetCellsForRow(Row row, IList<string> columnLetters)
        {
            var workIdx = 0;
            var emptyCell = new Cell { DataType = null, CellValue = new CellValue(string.Empty) };
            foreach (var cell in row.Descendants<Cell>())
            {
                var cellLetter = GetColumnAddress(cell.CellReference);
                var currentActualIdx = columnLetters.IndexOf(cellLetter);
                if (currentActualIdx < 0)
                {
                    break;
                }
                for (; workIdx < currentActualIdx; workIdx++)
                {
                    yield return emptyCell;
                }
                
                yield return cell;
                workIdx++;

                if (cell != row.LastChild) continue;
                for (; workIdx < columnLetters.Count(); workIdx++)
                {
                    yield return emptyCell;
                }
            }                
        }

        public static void RemoveRow(Row row)
        {
            row.Remove();
        }
        
        private static string GetColumnAddress(string cellReference)
        {
            var regex = new Regex("[A-Za-z]+");
            var match = regex.Match(cellReference);
            return match.Value;
        }
    }
}
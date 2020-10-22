using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ExcelParsingUtil.Models;
using Microsoft.AspNetCore.Mvc.ModelBinding;

namespace ExcelParsingUtil
{
    public static class ImportService
    {
        public static async Task<ImportResult> ImportFile(ImportFileData importData, ModelStateDictionary modelState)
        {
            if ((importData.FileData.Length / 1024.0F / 1024.0F) > 1.1f)
            {
                AddModelError(
                    modelState,
                    "File exceeds maximum size of 1MB",
                    0
                );
                return null;
            }

            #region Variable Initialization

            IList<ImportRow> rows;
            IList<ComicBookInventory> newComicBookInventory = new List<ComicBookInventory>();
            var expectedColumns = new[]
            {
                "issuenumber",
                "title",
                "description",
                "flag",
                "datepublished",
            };

            #endregion

            using (var document = ExcelParsingUtil.OpenDocumentFromByteArray(importData.FileData, true))
            {
                var worksheetPart = ExcelParsingUtil.GetWorksheetPartFromDocument(document, "");
                if (!ExcelParsingUtil.HasAnyRows(worksheetPart))
                {
                    AddModelError(
                        modelState,
                        "File is empty",
                        0
                    );
                    return null;
                }

                #region Data Converter

                ImportRow DataConverter(IList<string> rowData, IList<string> columnNames,
                    int rowNumber)
                {
                    if (expectedColumns.Except(columnNames).Any())
                    {
                        var exception = new Exception("Missing expected columns in data");
                        exception.Data.Add("rowNumber", rowNumber);
                        throw exception;
                    }

                    var importRow = new ImportRow
                    {
                        IssueNumber = rowData[columnNames.IndexOf("issuenumber")].Trim(),
                        Title = rowData[columnNames.IndexOf("title")].Trim(),
                        Description = rowData[columnNames.IndexOf("description")].Trim(),
                        Flag = rowData[columnNames.IndexOf("flag")].Trim(),
                        DatePublished = DateTime.FromOADate(double.Parse(rowData[columnNames.IndexOf("datepublished")].Trim())),
                        ItemGuid = Guid.NewGuid(),
                        RowNumber = rowNumber,
                    };

                    ValidateImportRow(importRow, modelState);

                    return importRow;
                }

                #endregion

                #region Parse File and Validate

                try
                {
                    rows = ExcelParsingUtil.GetDataToList(
                        document,
                        DataConverter
                    );
                }
                catch (Exception e)
                {
                    var rowNumber = 0;
                    if (e.Data.Keys.Count > 0)
                    {
                        int.TryParse(e.Data["rowNumber"].ToString(), out rowNumber);
                    }
                    AddModelError(
                        modelState,
                        e.Message,
                        rowNumber
                    );
                    return null;
                }

                if (!rows.Any())
                {
                    AddModelError(
                        modelState,
                        "No rows",
                        0
                    );
                    return null;
                }

                if (TooManyErrors(modelState))
                {
                    return null;
                }

                #endregion

                #region Loop Rows for each Sample Result for Review

                foreach (var row in rows)
                {
                    if (row == null)
                    {
                        continue;
                    }

                    try
                    {
                        ComicBookInventory currentComicBookInventory = null;
                        currentComicBookInventory = ProcessOneComicBook(row);
                        if (modelState.ErrorCount == 0)
                        {
                            newComicBookInventory.Add(currentComicBookInventory);
                        }
                    }
                    catch (Exception e)
                    {
                        AddModelError(
                            modelState,
                            e.Message,
                            row.RowNumber
                        );
                    }

                    if (!TooManyErrors(modelState))
                    {
                        continue;
                    }
                    break;
                }

                if (modelState.ErrorCount > 0)
                {
                    return null;
                }

                #endregion

                #region Persist Sample Results for Review objects

                // Using a repository save these in a database
                // await repository.AddMultiple(newSampleResultReview);
                // await DbContext.SaveChangesAsync();

                #endregion
            }

            #region Create Import Stats for Front-end

            var importStats = new ImportResult
            {
                NumberOfRows = rows.Count,
                FileName = importData.FileName
            };

            #endregion

            return importStats;
        }

        private static void AddModelError(ModelStateDictionary modelState, string message, int rowNumber)
        {
            var prefixMessage = "Error while processing row " + rowNumber + ": ";
            if (rowNumber == 0)
            {
                prefixMessage = "";
            }
            modelState.AddModelError("form", prefixMessage + message);
        }

        private static ComicBookInventory ProcessOneComicBook(
            ImportRow row
        )
        {
            return new ComicBookInventory()
            {
                // Id = row.ItemGuid,
                IssueNumber = row.IssueNumber,
                Title = row.Title,
                Description = row.Description,
                Flag = row.Flag,
                DatePublished = row.DatePublished,
            };
        }

        private static bool TooManyErrors(ModelStateDictionary modelState, bool addTooManyErrorsMessage = true)
        {
            var tooManyErrors = modelState?["form"]?.Errors?.Count > 10;
            if (tooManyErrors && addTooManyErrorsMessage)
            {
                AddModelError(
                    modelState,
                    "Stopped processing after 10 or more errors",
                    0
                );
            }

            return tooManyErrors;
        }

        #region Validation

        private static void ValidateImportRow(ImportRow row, ModelStateDictionary modelState)
        {
            if (string.IsNullOrEmpty(row.IssueNumber))
            {
                AddModelError(
                    modelState,
                    "Issue Number cannot be left blank",
                    row.RowNumber
                );
            }

            if (string.IsNullOrEmpty(row.Title))
            {
                AddModelError(
                    modelState,
                    "Title cannot be left blank",
                    row.RowNumber
                );
            }

            if (string.IsNullOrEmpty(row.Description))
            {
                AddModelError(
                    modelState,
                    "Description cannot be left blank",
                    row.RowNumber
                );
            }

            if (row.DatePublished > DateTimeOffset.Now)
            {
                AddModelError(
                    modelState,
                    "Date Published must be a past date",
                    row.RowNumber
                );
            }
        }

        #endregion
    }
}

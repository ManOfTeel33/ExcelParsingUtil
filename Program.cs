using System;
using ExcelParsingUtil.Models;
using Microsoft.AspNetCore.Mvc.ModelBinding;

namespace ExcelParsingUtil
{
    public static class ExcelParse
    {
        public static void Main(string[] args)
         {
             Console.WriteLine("Starting Retrieving data from Excel...");
             Console.WriteLine("");

             // Grab file and convert to byte array
             string filePath = @"C:\TestFile.xlsx";
             byte[] bytes = System.IO.File.ReadAllBytes(filePath);

             var importFileData = new ImportFileData
             {
                 FileData = bytes,
                 FileName = "Test File"
             };
             var modelState = new ModelStateDictionary();

            var import = ImportService.ImportFile(importFileData, modelState);

             Console.WriteLine(import);

             Console.WriteLine("");
             Console.WriteLine("Press any key to exit...");
             Console.ReadKey();
         }
    }
}

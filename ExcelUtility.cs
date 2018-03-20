using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using SpreadsheetLight;
using System.Drawing;

namespace Lazy.Utility.DataPreprocessor
{
  public static class ExcelUtility
  {
    /// <summary>
    /// Access renderedSheet by renderedSheets[$"{realTableName}.{realSheetName}"]
    /// </summary>
    /// <param name="excelFilePaths"></param>
    /// <returns></returns>
    public static Dictionary<string, RenderedSheet> RenderExcelSheets(string[] excelFilePaths)
    {
      var renderedSheets = new RenderedSheetDictionary();
      foreach (var tablePath in excelFilePaths)
      {
        var tableName = Path.GetFileNameWithoutExtension(tablePath);
        UseSpreadsheetLight(tablePath, (sl, sheetName) => {
          sl.SelectWorksheet(sheetName);
          var cells = sl.GetCells(); // cells[Y][X] | cells[ID][Property]
          var rowCount = cells.Count;
          var colCount = cells[1].Count;

          var unrenderedSheet = new Dictionary<string, RenderedRow>();
          for (int rowIndex = 2; rowIndex <= rowCount; rowIndex++) // SpreadsheetLight rowIndex start with 1, here we start at 2 for skip caption row.
          {
            var id = sl.GetCellValueAsString(rowIndex, 1);
            if (id == "") continue;

            var row = new RenderedRow();
            for (int colIndex = 1; colIndex <= colCount; colIndex++) // SpreadsheetLight colIndex start with 1
            {
              var cellCaption = sl.GetCellValueAsString(1, colIndex);
              var cellValue = sl.GetCellValueAsString(rowIndex, colIndex);
              var cellColor = sl.GetCellStyle(rowIndex, colIndex).Fill.PatternForegroundColor.Name;

              row.Add(cellCaption, new RenderedCell {
                value = cellValue,
                color = cellColor
              });
            }
            unrenderedSheet.Add(id, row);
          }
          renderedSheets.AddRenderedSheet(new RenderedSheet(tableName, sheetName, unrenderedSheet));
        });
      }
      return renderedSheets;
    }

    public static void UseSpreadsheetLight(string excelFilePath, Action<SLDocument, string> action)
    {
      using (var fs = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
      using (var sl = new SLDocument(fs))
      {
        var sheetNames = sl.GetSheetNames();
        foreach (var sheetName in sheetNames)
        {
          action(sl, sheetName);
        }
      }
    }
  }
}

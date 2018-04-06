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
    public static Dictionary<string, RenderedSheet> RenderExcelSheets(IEnumerable<string> excelFilePaths)
    {
      var renderedSheets = new RenderedSheetDictionary();
      foreach (var tablePath in excelFilePaths)
      {
        var tableName = Path.GetFileNameWithoutExtension(tablePath);
        UseSpreadsheetLight(tablePath, (sl, sheetName) => {
          var tableMode = sheetName[0] == '!'; // tableMode: from (1,1) to (max,max)

          sl.SelectWorksheet(sheetName);
          var cells = sl.GetCells(); // cells[Y][X] | cells[ID][Property]
          var rowCount = cells.Count;
          int colCount;
          try
          {
            colCount = cells[1].Count;
          }
          catch (KeyNotFoundException ex)
          {
            colCount = 0;
          }

          var unrenderedSheet = new Dictionary<string, RenderedRow>();

          if (!tableMode)
          {
            for (int rowIndex = 2; rowIndex <= rowCount; rowIndex++) // SpreadsheetLight rowIndex start with 1, here we start at 2 for skip caption row.
            {
              var id = sl.GetCellValueAsString(rowIndex, 1);
              if (id == "") continue;

              int n;
              var result = int.TryParse(id, out n);
              if (result == false)
              {
                throw new InvalidDataException($"无法正确解析普通表格{tableName}（表单：{sheetName}）中的ID：\"{id}\"，是否顶部有空行？或者在ID列中输入了非整数的值？");
              }

              var row = new RenderedRow();
              for (int colIndex = 1; colIndex <= colCount; colIndex++) // SpreadsheetLight colIndex start with 1
              {
                var cellCaption = sl.GetCellValueAsString(1, colIndex);
                var cellValue = sl.GetCellValueAsString(rowIndex, colIndex);
                var cellColor = sl.GetCellStyle(rowIndex, colIndex).Fill.PatternForegroundColor.Name;

                try
                {
                  row.Add(cellCaption, new RenderedCell {
                    text = cellValue,
                    color = cellColor
                  });
                }
                catch (ArgumentException ex)
                {
                  Console.WriteLine($"读取普通表格（{tableName} - {sheetName}）时发现异常：{ex.Message}");
                  Console.ReadKey();
                  throw;
                }
              }
              unrenderedSheet.Add(id, row);
            }
          }
          else
          {
            for (int rowIndex = 1; rowIndex <= rowCount; rowIndex++) // SpreadsheetLight rowIndex start with 1, here we start at 1 for capture all row.
            {
              var row = new RenderedRow();
              for (int colIndex = 1; colIndex <= colCount; colIndex++) // SpreadsheetLight colIndex start with 1
              {
                var cellValue = sl.GetCellValueAsString(rowIndex, colIndex);
                var cellColor = sl.GetCellStyle(rowIndex, colIndex).Fill.PatternForegroundColor.Name;

                row.Add(colIndex.ToString(), new RenderedCell {
                  text = cellValue,
                  color = cellColor
                });
              }
              unrenderedSheet.Add(rowIndex.ToString(), row);
            }
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

    public static void UseSpreadsheetLight(string excelFilePath, string excelSheetName, Action<SLDocument> action)
    {
      UseSpreadsheetLight(excelFilePath, (sl, sheetName) => {
        if (sheetName == excelSheetName)
        {
          sl.SelectWorksheet(sheetName);
          action(sl);
        }
      });
    }
  }
}

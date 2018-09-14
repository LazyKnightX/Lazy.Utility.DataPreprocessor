using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using SpreadsheetLight;
using System.Drawing;
using System.Runtime.InteropServices;
using log4net;

namespace Lazy.Utility.DataPreprocessor
{
  public static class ExcelUtility
  {
    private static ILog logger => ExcelDataPreprocessor.logger;
    /// <summary>
    /// Access renderedSheet by renderedSheets[$"{realTableName}.{realSheetName}"]
    /// </summary>
    /// <param name="excelFilePaths"></param>
    /// <param name="useTableModeSheets"></param>
    /// <returns></returns>
    public static Dictionary<string, RenderedSheet> RenderExcelSheets(IEnumerable<string> excelFilePaths, IEnumerable<string> useTableModeSheets = null)
    {
      var renderedSheets = new RenderedSheetDictionary();
      foreach (var tablePath in excelFilePaths)
      {
        var tableName = Path.GetFileNameWithoutExtension(tablePath);
        logger.InfoFormat("正在提取表格：{0}（{1}）", tableName, tablePath);
        UseSpreadsheetLight(tablePath, (sl, sheetName) => {
          var useTableMode = useTableModeSheets != null && useTableModeSheets.Contains(sheetName);

          sl.SelectWorksheet(sheetName);
          var cells = sl.GetCells(); // cells[Y][X] | cells[ID][Property]
          int rowCount;
          try
          {
            rowCount = cells.Keys.Max();
          }
          catch
          {
            rowCount = 0;
          }

          var unrenderedSheet = new Dictionary<string, RenderedRow>();

          // 正常的数据表模式（每行一条的数据）
          if (!useTableMode)
          {
            // SpreadsheetLight的rowIndex从1开始，这里设置为2，从而跳过标题行
            const int firstRowValueIndex = 2;
            for (int rowIndex = firstRowValueIndex; rowIndex <= rowCount; rowIndex++)
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
              var colCount = cells[rowIndex].Count;

              // SpreadsheetLight colIndex start with 1
              const int firstColValueIndex = 1;
              for (int colIndex = firstColValueIndex; colIndex <= colCount; colIndex++)
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
                  Console.WriteLine($"提取数据表（{tableName} - {sheetName}）时发现异常：{ex.Message}");
                  Console.ReadKey();
                  throw;
                }
              }
              unrenderedSheet.Add(id, row);
            }
          }
          // 源表格模式（从(1,1)到(max,max)），通常用于特殊编译的规则
          else
          {
            foreach (KeyValuePair<int,Dictionary<int,SLCell>> pair in cells)
            {
              var rowIndex = pair.Key;
              var colValues = pair.Value;
              var colCount = colValues.Keys.Max();

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
        logger.InfoFormat("结束提取表格：{0}（{1}）", tableName, tablePath);
        // TODO 低优先,不急 优化ExcelUtility的效率
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

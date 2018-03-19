using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using SpreadsheetLight;

namespace Lazy.Utility.DataPreprocessor
{
  public static class ExcelUtility
  {
    /// <summary>
    /// Access renderedSheet by renderedSheets[$"{realTableName}.{realSheetName}"]
    /// </summary>
    /// <param name="excelFilePaths"></param>
    /// <param name="ignoredSheetNamePrefixes">sheet name start with those prefix (to "char") will be ignored.</param>
    /// <returns></returns>
    public static Dictionary<string, RenderedSheet> RenderExcelSheets(string[] excelFilePaths, string ignoredSheetNamePrefixes = "@~")
    {
      var renderedSheets = new RenderedSheetDictionary();
      foreach (var tablePath in excelFilePaths)
      {
        var realTableName = Path.GetFileNameWithoutExtension(tablePath);
        Dictionary<string, object[][]> excelTable = ReadAll(tablePath);
        foreach (var excelSheet in excelTable)
        {
          var realSheetName = excelSheet.Key;
          var realSheetData = excelSheet.Value;
          if (ignoredSheetNamePrefixes.Contains(realSheetName[0])) continue;
          renderedSheets.AddRenderedSheet(RenderSheet(realTableName, realSheetName, realSheetData));
        }
      }
      return renderedSheets;
    }

    private static void UseSpreadsheetLight(string excelFilePath, string excelSheetName, Action<SLDocument> action)
    {
      using (var fs = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
      using (var sl = new SLDocument(fs, excelSheetName))
      {
        action(sl);
      }
    }

    /// <summary>
    /// https://www.connectionstrings.com
    /// "HDR=No" means "Do not ignore caption line, or the result will not contain caption line in dt.Rows"
    /// </summary>
    private static string MakeConnectionString(string pathName)
    {
      string connectionString = $"Data Source={pathName};";

      FileInfo file = new FileInfo(pathName);
      if (!file.Exists) { throw new FileNotFoundException(pathName); }

      switch (file.Extension)
      {
        // https://www.connectionstrings.com
        // "HDR=No" means "Do not ignore caption line, or the result will not contain caption line in dt.Rows"
        case ".xls":
          connectionString += "Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties='Excel 8.0;HDR=No;IMEX=1'";
          break;
        case ".xlsx":
          connectionString += "Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties='Excel 12.0;HDR=No;IMEX=1'";
          break;
        default:
          connectionString += "Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties='Text;HDR=No;IMEX=1'";
          break;
      }

      return connectionString;
    }

    private static Dictionary<string, Dictionary<int, Dictionary<int, string>>> ReadAll(string filePath)
    {
      var sheets = new Dictionary<string, Dictionary<int, Dictionary<int, string>>>();

      string strConn = MakeConnectionString(filePath);
      DataTable dt = new DataTable();

      OleDbConnection conn = new OleDbConnection(strConn);
      conn.Open();

      DataTable excel = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
      if (excel == null)
      {
        throw new NullReferenceException($"Invalid Excel with Path: {filePath}");
      }
      foreach (DataRow sheetInfo in excel.Rows)
      {
        string sheetName = sheetInfo["TABLE_NAME"].ToString().Replace("'", "");

        dt.Clear();
        new OleDbDataAdapter($"SELECT * FROM [{sheetName}]", conn).Fill(dt);
        string sheetClearName = sheetName.Remove(sheetName.Length - 1, 1); // Remove last "$"

        var output = new Dictionary<int, Dictionary<int, string>>();

        for (int y = 0; y < dt.Rows.Count; y++)
        {
          DataRow row = dt.Rows[y];

          output[y] = new Dictionary<int, string>();
          for (int i = 0; i < row.ItemArray.Length; i++)
          {
            output[y].Add(i, (string)row.ItemArray[i]);
          }
        }
        sheets.Add(sheetClearName, output);
      }

      return sheets;
    }

    private static RenderedSheet RenderSheet(string fileName, string sheetName, object[][] sheetTable)
    {
      var container = new Dictionary<string, RenderedRow>();

      // 行数
      var rowMax = sheetTable.Length;

      // 列数
      var colMax = sheetTable[0].Length;

      // 行左侧的标题
      var rowsCaption = new Dictionary<int, string>();
      for (var row = 0; row < rowMax; row++) rowsCaption.Add(row, sheetTable[row][0].ToString());

      // 列顶部的标题
      var colsCaption = new Dictionary<int, string>();
      for (var col = 0; col < colMax; col++) colsCaption.Add(col, sheetTable[0][col].ToString());

      // 这里不用在意左上单元格的标题，默认永远为“ID”
      for (var row = 1; row < rowMax; row++) // 从1开始是为了跳过行左侧的标题
      {
        for (var col = 1; col < colMax; col++) // 从1开始是为了跳过列顶部的标题
        {
          var rowCaption = rowsCaption[row]; // 行标题
          var colCaption = colsCaption[col]; // 列标题
          var cellValue = sheetTable[row][col].ToString(); // 单元格内容

          // 忽略没有ID 或 没有属性标记 或 空白 的数据格
          if (rowCaption == String.Empty || colCaption == String.Empty || cellValue == String.Empty) continue;

          // 如果是第一次读取“行”，先初始化这行
          if (!container.ContainsKey(rowCaption)) container.Add(rowCaption, new RenderedRow());

          // 记录单元格
          container[rowCaption].Add(colCaption, new RenderedCell {
            value = cellValue,
            bgColor = cellBgColor
          });
          /*
          try
          {
            // 记录单元格
            container[rowCaption].Add(colCaption, cellValue);
          }
          catch (ArgumentException e)
          {
            logger.Error($@"表格错误：某个标题列出现重复值。

文件名字：{fileName}
表格名字：{sheetName}

纵列顶部标题：{colCaption}
横行左侧标题：{rowCaption}
单元格数据：{cellValue}", e);
            Console.ReadKey();

            return null;
          }
          */
        }
      }

      return new RenderedSheet(fileName, sheetName, container);
    }
  }
}

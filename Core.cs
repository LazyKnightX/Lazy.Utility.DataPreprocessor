using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Xml;
using log4net;
using Newtonsoft.Json;
using System.Windows.Forms;
using SpreadsheetLight;

namespace Lazy.Utility.DataPreprocessor.Core
{
  public static class Extends
  {
    public static string[] Split(this string self, string separator, StringSplitOptions options = StringSplitOptions.RemoveEmptyEntries)
    {
      return self.Split(new[] {separator}, options);
    }

    public static string[] Split(this string self, string[] separator)
    {
      return self.Split(separator, StringSplitOptions.RemoveEmptyEntries);
    }
  }

  public class Config
  {
    private string path;
    private XmlElement root;
    public Dictionary<string, string> sheets;
    public Dictionary<string, string> properties;

    public Config(string path)
    {
      XmlDocument doc;
      this.path = path;

      XmlReaderSettings readerSettings = new XmlReaderSettings {
        IgnoreWhitespace = true,
        IgnoreComments = true,
        CheckCharacters = true,
        CloseInput = true,
        IgnoreProcessingInstructions = false,
        ValidationFlags = System.Xml.Schema.XmlSchemaValidationFlags.None,
        ValidationType = ValidationType.None
      };

      using (XmlReader reader = XmlReader.Create(path, readerSettings))
      {
        doc = new XmlDocument();
        doc.Load(reader);
      }

      root = doc.DocumentElement;
      _InitSheets();
      _InitProperties();
    }

    private void _InitSheets()
    {
      sheets = new Dictionary<string, string>();

      var nodes = root.SelectNodes("//sheet");
      if (nodes == null) throw new NullReferenceException($"{path}中缺少sheet。");

      foreach (XmlNode node in nodes)
      {
        var attrs = node.Attributes;
        if (attrs == null) throw new NullReferenceException($"{path}中的sheet必须包含字段name和input。");
        var name = attrs["name"];
        if (attrs["name"] == null) throw new NullReferenceException($"{path}中的sheet必须包含字段name。");
        var input = attrs["input"];
        if (attrs["input"] == null) throw new NullReferenceException($"{path}中的sheet必须包含字段input。");
        sheets.Add(name.InnerText, input.InnerText);
      }
    }

    private void _InitProperties()
    {
      properties = new Dictionary<string, string>();

      var nodes = root.SelectNodes($"//property");
      if (nodes == null) throw new NullReferenceException($"{path}中缺少property。");

      foreach (XmlNode node in nodes)
      {
        var attrs = node.Attributes;
        if (attrs == null) throw new NullReferenceException($"{path}中的property必须包含字段name和value。");
        var name = attrs["name"].Value;
        if (name == null) throw new NullReferenceException($"{path}中的property必须包含字段name。");
        var value = attrs["value"].Value;
        if (value == null) throw new NullReferenceException($"{path}中的property必须包含字段value。");

        properties.Add(name, value);
      }
    }
  }

  /// <summary>
  /// DO NOT USE OBJECT INITIALIZER FOR BETTER DEBUG EXPERIENCE.
  /// </summary>
  public abstract class ExcelDataPreprocessor
  {
    protected abstract void Compile();
    protected abstract void Output();

    private string configPath = "config.xml";
    protected Config config;
    protected Dictionary<string, RenderedSheet> renderedSheets;
    protected static readonly ILog logger = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);
    protected string[] itemSeparators = {"\n", "、"};

    public void Execute()
    {
      logger.Info("开始转表，肝肝肝肝肝肝!");
      _RegisterAllExceptionCatcher();
      logger.Info("正在读取设置。");
      _LoadConfig();
      logger.Info("正在读取表格。");
      _RenderExcelSheets();
      logger.Info("正在编译表格。");
      Compile();
      logger.Info("正在输出Json。");
      Output();
      logger.Info("全部搞定了！收工领盒饭！");
      Console.ReadKey();
    }

    private void _RegisterAllExceptionCatcher()
    {
      Application.ThreadException += (sender, e) => {
        logger.Fatal("Unhandled Thread Exception Catched !!!", e.Exception);
        Console.ReadKey();
      };
      AppDomain.CurrentDomain.UnhandledException += (sender, e) => {
        logger.Fatal("Unhandled Exception Catched !!!", e.ExceptionObject as Exception);
        Console.ReadKey();
      };
    }

    private void _LoadConfig()
    {
      logger.Info($"正在加载配置…… - {configPath}");
      try
      {
        config = new Config(configPath);
      }
      catch
      {
        logger.Error($"无法加载配置文件，本工具目录下的config.xml是否丢失了？\n路径：{Path.GetFullPath(configPath)}");
        Console.ReadKey();
      }
    }

    private void _RenderExcelSheets()
    {
      try
      {
        renderedSheets = ExcelUtility.RenderExcelSheets(config.sheets.Values.ToArray());
      }
      catch (FileNotFoundException e)
      {
        logger.Error("没有找到数据表文件， config.xml 中的数据表的路径配置正确了吗？", e);
        Console.ReadKey();
      }
      catch (InvalidOperationException e)
      {
        // https://xywiki.com/p/System.Data.OleDb
        logger.Error("未安装 Microsoft Access Database Engine 或者需要安装64位的 Microsoft Access Database Engine 。", e);
        Console.ReadKey();
      }
      catch (Exception e)
      {
        logger.Error("未知错误。", e);
        Console.ReadKey();
      }
    }

    protected float CompilePercent(string text)
    {
      ValidEmptyData("<百分比格子>", text);
      return text.Contains("%") ? Convert.ToSingle(text.Replace("%", "")) / 100 : Convert.ToSingle(text);
    }

    /// <summary>
    /// will auto multiply by 100.
    /// </summary>
    /// <param name="text"></param>
    /// <returns></returns>
    protected int CompilePercent100(string text)
    {
      ValidEmptyData("<百分比格子>", text);
      return IsIgnoredCell(text) ? 0 : Convert.ToInt32(CompilePercent(text) * 100);
    }

    protected int CompileInt32(string text)
    {
      ValidEmptyData("<整数格子>", text);
      return IsIgnoredCell(text) ? 0 : Convert.ToInt32(text);
    }

    protected double CompileDouble(string text)
    {
      ValidEmptyData("<实数>", text);
      return Convert.ToDouble(text);
    }

    protected void ValidEmptyData(string caption, string value, string hint = null)
    {
      if (!string.IsNullOrEmpty(value)) return;
      if (hint == null)
        logger.Warn($"检测到空【{caption}】，忘了填写数据？");
      else
        logger.Warn($"检测到空【{caption}】（{hint}），忘了填写数据？");
      Console.ReadKey();
    }

    protected void ValidGroupData(string name, string name2, string[] valueGroup, string[] valueGroup2)
    {
      if (valueGroup.Length != valueGroup2.Length)
      {
        logger.Warn($"检测到没有对齐的Group数据组【{name}】【{name2}】：\"{string.Join("、", valueGroup)}\"、\"{string.Join("、", valueGroup2)}\"，填漏了？");
        Console.ReadKey();
      }
    }

    protected bool IsIgnoredCell(string value)
    {
      return value == "/";
    }

    protected void UpdateData<TData>(Dataset<TData> dataset, int id, TData data)
    {
      if (dataset.ContainsKey(id))
      {
        logger.Info("更新ID：" + id);
        dataset[id] = data;
      }
      else
      {
        logger.Info("添加ID：" + id);
        dataset.Add(id, data);
      }
    }

    /// <summary>
    ///
    /// </summary>
    /// <param name="jsonFileName"></param>
    /// <param name="dataset"></param>
    /// <param name="pushEmpty">rpg maker mv need first cell as empty.</param>
    /// <typeparam name="T"></typeparam>
    protected void OutputJson<T>(string jsonFileName, Dataset<T> dataset, bool pushEmpty)
    {
      var list = pushEmpty ? new List<T> {default} : new List<T>();
      foreach (var data in dataset) list.Add(data.Value);
      var output = JsonConvert.SerializeObject(list, Newtonsoft.Json.Formatting.Indented);
      File.WriteAllText(GetJsonOutputPath(jsonFileName), output);
      logger.Info($"已成功输出：{jsonFileName}。");
    }

    private string GetJsonOutputPath(string jsonFileName)
    {
      return String.Format(config.properties["JsonFileOutputPathTemplate"], jsonFileName);
    }

    /// <summary>
    /// dataType: Item, Weapon, Armor, Skill, ...
    /// sheetType: Normal, Feature, ...
    /// * THEY ARE ALL DEFINED BY "config.xml".
    /// </summary>
    /// <param name="tableName"></param>
    /// <param name="sheetName"></param>
    /// <returns></returns>
    private RenderedSheet GetRenderedSheet(string tableName, string sheetName)
    {
      try
      {
        return renderedSheets[$"{tableName}.{sheetName}"];
      }
      catch (KeyNotFoundException ex)
      {
        logger.Error($"没有找到表格 {tableName} 和 表单 {sheetName} ，是否忘记了注册需要被加载的表格的文件地址？", ex);
      }
      return null;
    }

    protected void LoopRenderedSheet(RenderedSheet sheet, Action<int, RenderedRow> action)
    {
      sheet.ForEachRow(action);
    }

    protected void LoopRenderedSheet(string tableName, string sheetName, Action<int, RenderedRow> action)
    {
      LoopRenderedSheet(GetRenderedSheet(tableName, sheetName), action);
    }
  }

  public class RenderedSheetDictionary : Dictionary<string, RenderedSheet>
  {
    public RenderedSheet GetRenderedSheet(string realTableName, string realSheetName)
    {
      return this[$"{realTableName}.{realSheetName}"];
    }

    public void AddRenderedSheet(RenderedSheet renderedSheet)
    {
      Add($"{renderedSheet.TableName}.{renderedSheet.SheetName}", renderedSheet);
    }
  }

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

  public class RenderedSheet
  {
    private Dictionary<string, RenderedRow> _dic;
    public string TableName;
    public string SheetName;

    public RenderedSheet(string tableName, string sheetName, Dictionary<string, RenderedRow> excelDataDictionary)
    {
      _dic = excelDataDictionary;
      TableName = tableName;
      SheetName = sheetName;
    }

    public RenderedCell this[string id, string property]
    {
      get
      {
        if (!_dic.ContainsKey(id))
        {
          throw new ArgumentNullException($"正在尝试通过一个没有被包含在数据表内的ID获取数据，或者忘记了在副表中添加对应的行？：（ID：{id}）");
        }
        var rows = _dic[id];

        if (!rows.ContainsKey(property))
        {
          throw new ArgumentNullException($"正在尝试通过一个没有被包含在数据表内的表头获取数据？或者表格内某个单元格有空值？：（ID: {id}, 表头: {property}）");
        }
        var value = rows[property];

        return value;
      }
    }

    public RenderedCell this[int id, string property] => this[id.ToString(), property];

    public RenderedRow this[string id]
    {
      get
      {
        if (!_dic.ContainsKey(id))
        {
          throw new ArgumentNullException($"正在尝试通过一个没有被包含在数据表内的ID获取数据，或者忘记了在副表中添加对应的行？：（ID：{id}）");
        }
        return _dic[id];
      }
    }

    public RenderedRow this[int id] => this[id.ToString()];

    public void ForEachRow(Action<int, RenderedRow> action)
    {
      foreach (var kv in _dic)
      {
        action(Convert.ToInt32(kv.Key), kv.Value);
      }
    }
  }

  public class Dataset<TValue> : Dictionary<int, TValue>
  {
  }

  public class RenderedRow
  {
    private Dictionary<string, RenderedCell> _dic = new Dictionary<string, RenderedCell>();
    public bool ContainsKey(string key) => _dic.ContainsKey(key);
    public void Add(string key, RenderedCell value) => _dic.Add(key, value);

    public RenderedCell this[string key]
    {
      get => ContainsKey(key) ? _dic[key] : default;
      set => _dic[key] = value;
    }
  }

  public class RenderedCell
  {
    public string value;
    public string bgColor;
  }
}

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using log4net;
using Newtonsoft.Json;

namespace Lazy.Utility.DataPreprocessor
{
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
}

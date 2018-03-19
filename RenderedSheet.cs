using System;
using System.Collections.Generic;

namespace Lazy.Utility.DataPreprocessor
{
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
}

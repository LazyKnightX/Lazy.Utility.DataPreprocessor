using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;

namespace Lazy.Utility.DataPreprocessor
{
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
        var input = attrs["input"];
        if (attrs["input"] == null) throw new NullReferenceException($"{path}中的sheet必须包含字段input。");
        var name = Path.GetFileNameWithoutExtension(input.InnerText);
        sheets.Add(name, input.InnerText);
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
}

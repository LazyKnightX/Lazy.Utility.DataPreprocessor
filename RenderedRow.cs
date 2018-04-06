using System.Collections.Generic;

namespace Lazy.Utility.DataPreprocessor
{
  public class RenderedRow
  {
    private List<RenderedCell> _list = new List<RenderedCell>();
    private Dictionary<string, RenderedCell> _dic = new Dictionary<string, RenderedCell>();

    public bool ContainsKey(string key)
    {
      return _dic.ContainsKey(key);
    }

    public void Add(string key, RenderedCell value)
    {
      _dic.Add(key, value);
      _list.Add(value);
    }

    public RenderedCell this[string key]
    {
      get => ContainsKey(key) ? _dic[key] : default;
      set => _dic[key] = value;
    }

    public RenderedCell this[int index] => _list[index];
  }
}

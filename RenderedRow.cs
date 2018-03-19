using System.Collections.Generic;

namespace Lazy.Utility.DataPreprocessor
{
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
}

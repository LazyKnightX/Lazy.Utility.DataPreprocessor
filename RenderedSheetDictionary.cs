using System.Collections.Generic;

namespace Lazy.Utility.DataPreprocessor
{
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
}

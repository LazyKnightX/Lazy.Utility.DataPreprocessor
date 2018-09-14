using System;

namespace Lazy.Utility.DataPreprocessor
{
  public class RenderedCell
  {
    /// <summary>
    /// excel text value
    /// </summary>
    public string text;
    /// <summary>
    /// excel background color (Foreground Pattern Color in SpreadsheetLight).
    /// </summary>
    public string color;

    public static implicit operator string (RenderedCell cell)
    {
      return cell != null ? cell.text : "";
    }
  }
}

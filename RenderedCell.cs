using System;

namespace Lazy.Utility.DataPreprocessor
{
  public class RenderedCell
  {
    /// <summary>
    /// excel text value
    /// </summary>
    public string value;
    /// <summary>
    /// excel background color (Foreground Pattern Color in SpreadsheetLight).
    /// </summary>
    public string color;

    public static implicit operator string (RenderedCell cell)
    {
      return cell.value;
    }
  }
}

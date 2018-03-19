using System;

namespace Lazy.Utility.DataPreprocessor
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
}

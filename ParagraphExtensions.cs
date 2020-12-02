using System.Windows;
using System.Windows.Documents;

namespace StyleSnooper
{
    internal static class ParagraphExtensions
    {
        public static void AddRun(this Paragraph paragraph, Style style, string s)
        {
            paragraph.Inlines.Add(new Run(s) { Style = style });
        }

        public static void AddLineBreak(this Paragraph paragraph)
        {
            paragraph.Inlines.Add(new LineBreak());
        }
    }
}
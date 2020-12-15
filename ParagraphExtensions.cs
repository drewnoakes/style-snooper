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

        public static void RemoveLastLineBreak(this Paragraph paragraph)
        {
            while (paragraph.Inlines.LastInline is not LineBreak)
                paragraph.Inlines.Remove(paragraph.Inlines.LastInline);
            paragraph.Inlines.Remove(paragraph.Inlines.LastInline);
        }
    }
}
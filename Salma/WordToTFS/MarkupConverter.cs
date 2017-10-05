using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WordToTFS
{
    public static class MarkupConverter
    {
        public static string CleanHeader(this string html)
        {
            return html.Substring(html.IndexOf("<html", StringComparison.Ordinal));
        }

        public static string SafeHtml(this string html)
        {
            var sb = new StringBuilder();
            sb.Append("<html>");
            sb.Append("<body>");
            sb.Append(html);
            sb.Append("</body>");
            sb.Append("</html>");
            return sb.ToString();
        }

        private static IEnumerable<string> GetPaths(this string html)
        {
            html = html.CleanHeader();

            var starts = new List<int>();
            for (var i = 0; i < html.Length; i++)
            {
                if (i >= html.Length - 8)
                {
                    break;
                }

                i = html.IndexOf(@"file:///", i, StringComparison.Ordinal);
                if (i == -1)
                {
                    break;
                }
                starts.Add(i);
            }

            var ends = starts.Select(start => html.IndexOf('"', start)).ToList();

            return starts.Select((t, i) => html.Substring(t, ends[i] - t)).ToList();
        }

        public static IDictionary<string, string> GetLinks(this string html)
        {
            return html.GetPaths().Distinct().ToDictionary(path => path, path => path.Replace(@"file:///", ""));
        }

        public static string ClearComments(this string html)
        {
            html = html.CleanHeader();

            var starts = new List<int>();
            for (var i = 0; i < html.Length; i++)
            {
                if (i >= html.Length - 4)
                {
                    break;
                }

                i = html.IndexOf(@"<!--", i, StringComparison.Ordinal);
                if (i == -1)
                {
                    break;
                }
                starts.Add(i);
            }

            var ends = starts.Select(start => html.IndexOf(@"-->", start, StringComparison.Ordinal) + 3).ToList();

            var content = new StringBuilder(html).ToString(); 
            //Enable cleaning mso styling
            content = starts.Select((t, i) => html.Substring(t, ends[i] - t)).Aggregate(content, (current, comment) => current.Replace(comment, ""));

            content = content.Replace(@"<![if !vml]>", "");
            content = content.Replace(@"<![endif]>", "");




            content = content.Substring(content.IndexOf("<body"));
            content = content.Substring(content.IndexOf(">") + 1);
            content = content.Remove(content.LastIndexOf("</body>"), content.Length - content.LastIndexOf("</body>"));


            //deleting index from description
            if (content.Contains("<div style='mso-element:comment-list'>"))
            {
                content = content.Remove(content.IndexOf("<div style='mso-element:comment-list'>"));
            }

            for (int i = 0; ; i++)
            {
                if (!content.Contains(">["))
                {
                    break;
                }
                //content = content.Remove(content.IndexOf(">[")+1, 5);
                content = content.Remove(content.IndexOf(">[") + 1, (content.IndexOf("]</a>")+1) - (content.IndexOf(">[") + 1));
            }
            return content.Trim();

        }
    }
}

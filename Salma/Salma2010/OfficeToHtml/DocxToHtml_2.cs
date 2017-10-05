using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Salma2010
{
    /// <summary>
    /// List type enum
    /// </summary>
    internal enum ListType
    {
        Bulleted,
        Numbered
    }

    internal sealed partial class DocxToHtml
    {
        /// <summary>
        /// Get html
        /// </summary>
        /// <param name="html">html</param>
        public string GetHtml(string html)
        {
            HtmlDocument wordDoc = new HtmlDocument();
            HtmlDocument tfsDoc = new HtmlDocument();

            try
            {
                wordDoc.LoadHtml(html);

                for (int i = 0; i < wordDoc.DocumentNode.ChildNodes.Count; i++)
                    if (wordDoc.DocumentNode.ChildNodes[i].NodeType == HtmlNodeType.Text)
                        wordDoc.DocumentNode.ChildNodes[i].Remove();

                DoHtml(wordDoc, tfsDoc);
            }
            catch
            {
                throw new Exception(GetFormatErrorMessage());
            }

            return tfsDoc.DocumentNode.InnerHtml;
        }

        /// <summary>
        /// Create html node
        /// </summary>
        /// <param name="inputDoc">input document</param>
        /// <param name="outputDoc">output document</param>
        private void DoHtml(HtmlDocument inputDoc, HtmlDocument outputDoc)
        {
            foreach (HtmlNode node in inputDoc.DocumentNode.ChildNodes)
                DoHtml(node, outputDoc.DocumentNode);
        }

        /// <summary>
        /// Create html node
        /// </summary>
        /// <param name="input">input node</param>
        /// <param name="output">output node</param>
        private void DoHtml(HtmlNode input, HtmlNode output)
        {
            if (input.NodeType != HtmlNodeType.Element)
                return;

            if (IsList(input))
            {
                if (IsNewList(input, output))
                    output.AppendChild(CreateListWithElement(input, output));
                else
                    AppendListWithElement(input, output);
            }
            else
            {
                if (input.OriginalName.Equals("table", StringComparison.InvariantCultureIgnoreCase))
                    output.AppendChild(CreateTableElement(input));
                else
                    output.AppendChild(CreateParapraphElement(input));
            }
        }

        /// <summary>
        /// Format error message
        /// </summary>
        private static string GetFormatErrorMessage()
        {
            string message = string.Empty;

            PropertyInfo[] properties = typeof(Properties.Resources).GetProperties(BindingFlags.Static | BindingFlags.NonPublic);

            PropertyInfo prop = properties.Select(p => p).Where(p => p.Name == "lblFormatCorruptedLabel").FirstOrDefault();

            if (prop != null)
                message = (string)prop.GetValue(null, null);

            return message;
        }

        /// <summary>
        /// Append list with element
        /// </summary>
        /// <param name="input">input node</param>f
        /// <param name="output">output node</param>
        private void AppendListWithElement(HtmlNode input, HtmlNode output)
        {
            string listNumber = GetListNumber(input);
            int listLevel = GetListLevel(input);

            HtmlNode listNode = output.ChildNodes.Where(node => Regex.IsMatch(
                node.GetAttributeValue("style", string.Empty), listNumber)).LastOrDefault();

            HtmlNode prevNode = listNode.Descendants().Where(node => Regex.IsMatch(
                node.GetAttributeValue("style", string.Empty), listNumber)).LastOrDefault();

            if (prevNode != null)
            {
                int prevListLevel = GetListLevel(prevNode);
                prevNode = prevNode.ParentNode;

                if (listLevel > prevListLevel)
                    prevNode.AppendChild(AddLevelElement(input, listLevel - prevListLevel));

                if (listLevel == prevListLevel)
                    prevNode.AppendChild(AddLevelElement(input, 0));

                if (listLevel < prevListLevel)
                {
                    while (prevListLevel != listLevel)
                    {
                        prevNode = prevNode.ParentNode;
                        prevListLevel--;
                    }

                    prevNode.AppendChild(AddLevelElement(input, 0));
                }
            }
        }

        /// <summary>
        /// Add level element
        /// </summary>
        /// <param name="input">input node</param>
        /// <param name="level">level</param>
        private HtmlNode AddLevelElement(HtmlNode input, int level)
        {
            HtmlNode temp;
            HtmlNode newNode;

            string listTag = GetListTag(input);
            string listNumber = GetListNumber(input);
            int listLevel = GetListLevel(input);

            if (level > 0)
            {
                level--;
                newNode = HtmlNode.CreateNode(string.Format("<{0}></{0}>", listTag));
                newNode.AppendChild(CreateListElement(input));
            }
            else
            {
                newNode = CreateListElement(input);
            }

            while (level > 0)
            {
                temp = HtmlNode.CreateNode(string.Format("<{0}></{0}>", listTag));
                temp.AppendChild(newNode);
                newNode = temp;
                level--;
            }

            return newNode;
        }

        /// <summary>
        /// Is new list
        /// </summary>
        /// <param name="input">input node</param>
        /// <param name="output">output node</param>
        private bool IsNewList(HtmlNode input, HtmlNode output)
        {
            string listNumber = GetListNumber(input);

            HtmlNode lastNode = output.LastChild;

            if (lastNode == null)
                return true;
            else
                if (!Regex.IsMatch(lastNode.GetAttributeValue("style", string.Empty), listNumber))
                    return true;

            return false;
        }

        /// <summary>
        /// Create new table element
        /// </summary>
        /// <param name="input">input node</param>
        private HtmlNode CreateTableElement(HtmlNode input)
        {
            for (int i = 0; i < input.ChildNodes.Count; i++)
                if (!input.ChildNodes[i].OriginalName.Equals("tr", StringComparison.InvariantCultureIgnoreCase))
                    input.ChildNodes[i].Remove();

            HtmlNode newTable = input.CloneNode(false);
            HtmlNode newRow = null;
            HtmlNode newColumn = null;

            foreach (HtmlNode tr in input.ChildNodes)
            {
                newRow = newTable.AppendChild(tr.CloneNode(false));

                for (int i = 0; i < tr.ChildNodes.Count; i++)
                    if (!tr.ChildNodes[i].OriginalName.Equals("td", StringComparison.InvariantCultureIgnoreCase))
                        tr.ChildNodes[i].Remove();

                foreach (HtmlNode td in tr.ChildNodes)
                {
                    newColumn = newRow.AppendChild(td.CloneNode(false));

                    HtmlDocument content = new HtmlDocument();

                    foreach (HtmlNode item in td.ChildNodes)
                        DoHtml(item, content.DocumentNode);

                    foreach (HtmlNode item in content.DocumentNode.ChildNodes)
                        newColumn.AppendChild(item);
                }
            }

            return newTable;
        }

        /// <summary>
        /// Create new paragraph element
        /// </summary>
        /// <param name="input">input node</param>
        private HtmlNode CreateParapraphElement(HtmlNode input)
        {
            HtmlNode newParagraph = null;

            int margin = GetMargin(input);

            newParagraph = HtmlNode.CreateNode(string.Format("<{0}></{0}>", "div"));
            newParagraph.AppendChildren(input.ChildNodes);

            return newParagraph;
        }

        /// <summary>
        /// Create new list element
        /// </summary>
        /// <param name="input">input node</param>
        private HtmlNode CreateListElement(HtmlNode input)
        {
            HtmlNode newParagraph = HtmlNode.CreateNode(string.Format("<{0}></{0}>", "li"));

            // set attribute
            string listNumber = GetListNumber(input);
            int listLevel = GetListLevel(input);

            newParagraph.SetAttributeValue("style", string.Format("{0} level{1}", listNumber, listLevel));
            newParagraph.AppendChildren(input.ChildNodes);

            return newParagraph;
        }
     
        /// <summary>
        /// Is list
        /// </summary>
        /// <param name="input">input node</param>
        private bool IsList(HtmlNode input)
        {
            if (input.NodeType == HtmlNodeType.Element)
                return Regex.IsMatch(input.GetAttributeValue("style", string.Empty).ToLower(),
                    "mso-list");

            return false;
        }

        /// <summary>
        /// Create list level element
        /// </summary>
        /// <param name="input">input node</param>
        /// <param name="output">output node</param>
        private HtmlNode CreateListWithElement(HtmlNode input, HtmlNode output)
        {
            string listNumber = GetListNumber(input);
            int listLevel = GetListLevel(input);
            string listTag = GetListTag(input);
            int margin = GetMargin(input);

            HtmlNode liNode = HtmlNode.CreateNode(string.Format("<{0}></{0}>", listTag));

            liNode.SetAttributeValue("style", string.Format("{0} level{1}", listNumber, listLevel));
            liNode.AppendChild(AddLevelElement(input, listLevel - 1));

            return liNode;
        }

        /// <summary>
        /// Get margin
        /// </summary>
        /// <param name="input">input node</param>
        private int GetMargin(HtmlNode input)
        {
            string regex = @"(margin-left\s*:\s*)(\d+)(\D)";
            int margin = 0;

            Match match = Regex.Match(input.GetAttributeValue("style", string.Empty).ToLower(),
                              regex);

            if (match.Success)
                Int32.TryParse(Regex.Replace(match.Value, regex, m => m.Groups[2].Value), out margin);

            return margin;
        }

        /// <summary>
        /// Get list html tag
        /// </summary>
        /// <param name="input">input node</param>
        private string GetListTag(HtmlNode input)
        {
            ListType type = GetListType(input);

            if (type == ListType.Bulleted)
                return "ul";

            return "ol";
        }

        /// <summary>
        /// Get list number
        /// </summary>
        /// <param name="input">input node</param>
        private string GetListNumber(HtmlNode input)
        {
            string regex = @"(mso-list)(\s*:\s*)(l\d+)(\s+)";

            Match match = Regex.Match(input.GetAttributeValue("style", string.Empty).ToLower(),
                regex);

            if (match.Success)
                return Regex.Replace(match.Value, regex, m =>
                    string.Format("{0}:{1}", m.Groups[1].Value, m.Groups[3].Value));

            return string.Empty;
        }

        /// <summary>
        /// Get list level
        /// </summary>
        /// <param name="input">input node</param>
        private int GetListLevel(HtmlNode input)
        {
            string regex = @"(\s+level)(\d+)(\s*)";
            int level = 0;

            Match match = Regex.Match(input.GetAttributeValue("style", string.Empty).ToLower(),
                regex);

            if (match.Success)
                Int32.TryParse(Regex.Replace(match.Value, regex,
                    m => m.Groups[2].Value), out level);

            return level;
        }

        /// <summary>
        /// Return list type - bulleted or numbered
        /// </summary>
        /// <param name="input">input node</param>
        private ListType GetListType(HtmlNode input)
        {
            string regex = @"(\s+level\d+\s+)(ul|ol)"; // @"(font-family\s*:\s*[\""\']?)([a-zA-Z\s]*)([\""\']?\s*;)";
            string marker = string.Empty;

            Match m = Regex.Match(input.GetAttributeValue("style", string.Empty).ToLower(),
                    regex);

            if (m.Success)
            {
                marker = Regex.Replace(m.Value, regex, match => match.Groups[2].Value);

                if (marker.Equals("ul"))
                    return ListType.Bulleted;
                else
                    return ListType.Numbered;
            }

            return ListType.Numbered;
        }
    }
}
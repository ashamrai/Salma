using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Salma2010
{
    /// http://dx-jo.com/blog/c-code-to-wirte-html-code-in-a-word-document-using-openxml-sdk/
    /// <summary>
    /// Html To docx
    /// </summary>
    internal sealed class HtmlToDocx
    {
        /// <summary>
        /// Available tags
        /// </summary>
        private string[] availableTags = {   "p", "div", "blockquote", "br", "ol", "ul", "li", "table", "tbody", "tr", "td", "span", "b", "u", "i", "a", "img" };

        private static volatile HtmlToDocx instance = null;
        private static object syncRoot = new Object();

        /// <summary>
        /// Instance
        /// </summary>
        public static HtmlToDocx Instance
        {
            get
            {
                if (instance == null)
                {
                    lock (syncRoot)
                    {
                        if (instance == null)
                            instance = new HtmlToDocx();
                    }
                }

                return instance;
            }
        }

        /// <summary>
        /// Convert html to docx - import
        /// </summary>
        /// <param name="control"></param>
        /// <param name="html"></param>
        public void ToDocxImport(Range range, string html, int wiID, string wiField)
        {
            int actionCount = 0;

            string ra = Convert.ToChar(0x000D).ToString() + Convert.ToChar(0x0007).ToString(); // \r\a
            string folderName = Path.GetTempPath() + Guid.NewGuid().ToString(); // temp folder
            string fileName = folderName + "\\temp.docx"; // temp file

            Globals.ThisAddIn.Application.ScreenUpdating = false; // false to update screen

            Microsoft.Office.Interop.Word.Document doc = null;

            try
            {
                CreateTempDocument(html, folderName, fileName);

                // get range to insert
                doc = Globals.ThisAddIn.Application.Documents.Open(fileName, Visible: false, ReadOnly: true);
                Range insert = doc.Content;
                PrepareRange(ref range, ref insert);

                // delete old control
                Globals.ThisAddIn.Application.ActiveDocument.Range(range.Start - 1, range.End + 1).Select();
                Globals.ThisAddIn.Application.Selection.ClearFormatting(); actionCount++;

                // insert range
                range.FormattedText = insert.FormattedText; actionCount++;
                range.LanguageID = WdLanguageID.wdNoProofing;

                foreach (InlineShape image in range.InlineShapes)
                    image.LinkFormat.SavePictureWithDocument = true;

                range.Select();
                ExtendSelection();
                Globals.ThisAddIn.Application.ActiveDocument.Range(range.End, range.End + 1).Delete(); actionCount++;
               
                // add new control
                Globals.ThisAddIn.AddWIControl(wiID, wiField, range);
            }
            catch(Exception e)
            {
                if (actionCount > 0)
                    Globals.ThisAddIn.Application.ActiveDocument.Undo(actionCount);

                throw new Exception(e.Message);
            }
            finally
            {
                Globals.ThisAddIn.Application.Selection.Collapse();
                Globals.ThisAddIn.Application.ScreenUpdating = true;

                try
                {
                    ((_Document)doc).Close(SaveChanges: false);

                    if (Directory.Exists(folderName))
                        Directory.Delete(folderName, true);
                }
                catch { }
            }
        }

        /// <summary>
        /// Convert html to docx
        /// </summary>
        /// <param name="control"></param>
        /// <param name="html"></param>
        public void ToDocx(ContentControl control, string html)
        {
            int actionCount = 0;

            string ra = Convert.ToChar(0x000D).ToString() + Convert.ToChar(0x0007).ToString(); // \r\a
            string folderName = Path.GetTempPath() + Guid.NewGuid().ToString(); // temp folder
            string fileName = folderName + "\\temp.docx"; // temp file
            string controlID = control.ID;

            Range range = control.Range;

            Globals.ThisAddIn.Application.ScreenUpdating = false; // false to update screen

            Microsoft.Office.Interop.Word.Document doc = null;

            try
            {
                // get wi id and field name
                string wiID = Globals.ThisAddIn.Application.ActiveDocument.Variables[control.ID + "_wiid"].Value;
                string wiField = Globals.ThisAddIn.Application.ActiveDocument.Variables[control.ID + "_wifield"].Value;
                string _ctrlId = control.ID;

                CreateTempDocument(html, folderName, fileName);

                // get range to insert
                doc = Globals.ThisAddIn.Application.Documents.Open(fileName, Visible: false, ReadOnly: true);
                Range insert = doc.Content;
                PrepareRange(ref range, ref insert);

                // delete old control

                Globals.ThisAddIn.Application.ActiveDocument.Range(range.Start - 1, range.End + 1).Select();
                if (wiField != "System.Title") Globals.ThisAddIn.Application.Selection.ClearFormatting(); actionCount++;
                
                control.Delete(true); actionCount++;

                // insert range
                if (wiField != "System.Title") { range.FormattedText = insert.FormattedText; actionCount++; }
                else { range.Text = insert.Text; actionCount++; }
                    
                range.LanguageID = WdLanguageID.wdNoProofing;

                foreach (InlineShape image in range.InlineShapes)
                    image.LinkFormat.SavePictureWithDocument = true;

                range.Select();
                ExtendSelection();

                Globals.ThisAddIn.Application.ActiveDocument.Range(range.End, range.End + 1).Delete(); actionCount++;

                // add new control
                int id = 0;

                if (Int32.TryParse(wiID, out id))
                    Globals.ThisAddIn.AddWIControl(id, wiField, range, _ctrlId);
            }
            catch (Exception e)
            {
                if (actionCount > 0)
                    Globals.ThisAddIn.Application.ActiveDocument.Undo(actionCount);

                throw new Exception(e.Message);
            }
            finally
            {
                Globals.ThisAddIn.Application.Selection.Collapse();
                Globals.ThisAddIn.Application.ScreenUpdating = true;

                try
                {
                    ((_Document)doc).Close(SaveChanges: false);

                    if (Directory.Exists(folderName))
                        Directory.Delete(folderName, true);
                }
                catch { }
            }

            Globals.ThisAddIn.DeleteVariables(controlID);
        }

        /// <summary>
        /// Extend selection
        /// </summary>
        /// <param name="range"></param>
        private void ExtendSelection()
        {
            string r = Convert.ToChar(0x000D).ToString(); // \r

            Microsoft.Office.Interop.Word.Range range = Globals.ThisAddIn.Application.ActiveDocument.Range(Globals.ThisAddIn.Application.Selection.Range.Start,
                Globals.ThisAddIn.Application.Selection.Range.End);

            if (range.Tables.Count > 0)
                foreach (Microsoft.Office.Interop.Word.Table table in range.Tables)
                {
                    if (range.Start > table.Range.Start)
                        range.Start = table.Range.Start;
                    if (range.End < table.Range.End)
                    {
                        range.End = table.Range.End; // +1;

                        Microsoft.Office.Interop.Word.Range nextChar = range.Next(Microsoft.Office.Interop.Word.WdUnits.wdCharacter);
                        if (nextChar != null && !nextChar.Text.Equals(r))
                            nextChar.InsertParagraphBefore();

                        range.End++;
                    }
                }

            if (range.Paragraphs.Count > 1)
                range = Globals.ThisAddIn.Application.ActiveDocument.Range(range.Paragraphs.First.Range.Start, range.Paragraphs.Last.Range.End);

            range.Select();
        }


        /// <summary>
        /// Create temp document
        /// </summary>
        /// <param name="html"></param>
        /// <param name="folderName"></param>
        /// <param name="fileName"></param>
        /// <returns></returns>
        private void CreateTempDocument(string html, string folderName, string fileName)
        {
            string bookmark = "insert";

            // create temporary document
            Microsoft.Office.Interop.Word.Document doc = Globals.ThisAddIn.Application.Documents.Add(
                Visible: false, DocumentType: WdDocumentType.wdTypeDocument);
            // add bookmark 
            doc.Bookmarks.Add(bookmark);

            // create temporary folder
            if (Directory.Exists(folderName))
                Directory.Delete(folderName, true);
            Directory.CreateDirectory(folderName);

            // save document 
            doc.SaveAs(fileName, AddToRecentFiles: false);
            ((_Document)doc).Close();

            // write html
            WriteDocx(bookmark, fileName, PrepareHtml(html, folderName));
        }

        /// <summary>
        /// Prepare range
        /// </summary>
        /// <param name="range"></param>
        /// <param name="insert"></param>
        private void PrepareRange(ref Range range, ref Range insert)
        {
            foreach (Microsoft.Office.Interop.Word.Paragraph paragraph in insert.Paragraphs)
                if (paragraph.Range.Text.Trim() == string.Empty && paragraph.Range.InlineShapes.Count == 0)
                    paragraph.Range.Delete();

            int paragraphLen = Globals.ThisAddIn.Application.ActiveDocument.Range(range.Paragraphs.First.Range.Start, range.Paragraphs.Last.Range.End).Characters.Count;
            int controlLen = Globals.ThisAddIn.Application.ActiveDocument.Range(range.Start - 1, range.End + 1).Characters.Count;

            // cannot copy range with paragraphs count > 1 to control in the middle of line
            if (paragraphLen - controlLen > 1 && insert.Paragraphs.Count > 1)
                throw new Exception();

            string r = Convert.ToChar(0x000D).ToString(); // \r
            string v = Convert.ToChar(0x000B).ToString(); // \v

            for (int i = 1; i <= insert.Characters.Count; i++)
                if (insert.Characters[i].Text.Equals(v))
                    insert.Characters[i].Text = r;
        }

        /// <summary>
        /// Prepare Html
        /// </summary>
        /// <param name="html"></param>
        /// <returns></returns>
        private string PrepareHtml(string html, string folderName)
        {
            HtmlAgilityPack.HtmlDocument input = new HtmlAgilityPack.HtmlDocument();
            HtmlAgilityPack.HtmlDocument output = new HtmlAgilityPack.HtmlDocument();

            input.LoadHtml(html);

            PrepareHtml(input.DocumentNode, output.DocumentNode);
            PrepareTables(output);
            PrepareImages(output, folderName);

            return output.DocumentNode.InnerHtml;
        }

        /// <summary>
        /// Prepare Html
        /// </summary>
        /// <param name="input"></param>
        /// <param name="output"></param>
        private void PrepareHtml(HtmlAgilityPack.HtmlNode input, HtmlAgilityPack.HtmlNode output)
        {
            HtmlAgilityPack.HtmlNode parent = output;

            switch (input.NodeType)
            {
                case HtmlAgilityPack.HtmlNodeType.Document:
                    break;
                case HtmlAgilityPack.HtmlNodeType.Element:

                    if (!availableTags.Contains(input.OriginalName))
                        return;

                    if (input.OriginalName.Equals("img"))
                    {
                        output.AppendChild(input.CloneNode(true));
                        return;
                    }

                    parent = output.AppendChild(input.CloneNode(false));

                    string newStyle = string.Empty;
                    string style = input.GetAttributeValue("style", string.Empty);

                    string href = input.GetAttributeValue("href", string.Empty);

                    parent.Attributes.RemoveAll();

                    if (style != string.Empty)
                        foreach (string item in style.Split(';'))
                            if ((item.Contains("font-weight") && item.Contains("bold")) ||
                                (item.Contains("font-style") && item.Contains("italic")) ||
                                (item.Contains("text-decoration") && item.Contains("underline")))
                                newStyle += string.Format("{0};", item);

                    if (newStyle != string.Empty)
                        parent.SetAttributeValue("style", newStyle);

                    if (href != string.Empty)
                        parent.SetAttributeValue("href", href);

                    break;
                case HtmlAgilityPack.HtmlNodeType.Text:
                    output.AppendChild(input.CloneNode(true));
                    return;
                default:
                    return;
            }

            foreach (HtmlAgilityPack.HtmlNode child in input.ChildNodes)
                PrepareHtml(child, parent);
        }

        /// <summary>
        /// Prepare tables
        /// </summary>
        /// <param name="html"></param>
        /// <returns></returns>
        private void PrepareTables(HtmlAgilityPack.HtmlDocument wordDoc)
        {
            List<HtmlAgilityPack.HtmlNode> tables = wordDoc.DocumentNode.Descendants().
               Where(node => node.OriginalName.Equals("table")).ToList<HtmlAgilityPack.HtmlNode>();

            for (int i = 0; i < tables.Count; i++)
            {
                HtmlAgilityPack.HtmlNode table = tables[i];

                table.Attributes.RemoveAll();
                table.SetAttributeValue("border", "1");
                table.SetAttributeValue("cellspacing", "0");
                table.SetAttributeValue("cellpadding", "0");
                table.SetAttributeValue("style", "border-collapse:collapse;border:none;");
                table.SetAttributeValue("width", "100%");

                List<HtmlAgilityPack.HtmlNode> rows = table.Descendants().
                    Where(node => node.OriginalName.Equals("tr")).ToList<HtmlAgilityPack.HtmlNode>();

                for (int j = 0; j < rows.Count; j++)
                {
                    HtmlAgilityPack.HtmlNode row = rows[j];
                    row.Attributes.RemoveAll();

                    List<HtmlAgilityPack.HtmlNode> columns = row.Descendants().
                                        Where(node => node.OriginalName.Equals("td")).ToList<HtmlAgilityPack.HtmlNode>();

                    for (int k = 0; k < columns.Count; k++)
                    {
                        HtmlAgilityPack.HtmlNode column = columns[k];
                        column.Attributes.RemoveAll();
                        column.SetAttributeValue("valign", "top");
                        string width = string.Format("{0}%", 100 / columns.Count);
                        column.SetAttributeValue("width", width);
                    }
                }
            }
        }

        /// <summary>
        /// Prepare images
        /// </summary>
        /// <param name="wordDoc"></param>
        /// <param name="folderName"></param>
        private void PrepareImages(HtmlAgilityPack.HtmlDocument wordDoc, string folderName)
        {
            List<HtmlAgilityPack.HtmlNode> images = wordDoc.DocumentNode.Descendants().
                Where(node => node.OriginalName.Equals("img")).ToList<HtmlAgilityPack.HtmlNode>();

            for (int i = 0; i < images.Count; i++)
            {
                HtmlAgilityPack.HtmlNode image = images[i];
                string src = image.GetAttributeValue("src", string.Empty);

                byte[] data;
                using (System.Net.WebClient client = new System.Net.WebClient())
                {
                    client.Credentials = System.Net.CredentialCache.DefaultCredentials;
                    data = client.DownloadData(src);
                }

                string imageSrc = folderName + "\\" + i + Path.GetExtension(src);
                File.WriteAllBytes(imageSrc, data);

                image.SetAttributeValue("src", imageSrc);
            }
        }

        /// <summary>
        /// Write docx
        /// </summary>
        /// <param name="bookMark"></param>
        /// <param name="filePath"></param>
        /// <param name="html"></param>
        private void WriteDocx(string bookMark, string filePath, string html)
        {
            try
            {
                string stringToWrite = "<html><head><meta charset=\"UTF-8\"></head><body>" + html + "</body></html>";

                MemoryStream memoryStream = new System.IO.MemoryStream(Encoding.UTF8.GetBytes(stringToWrite));

                using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
                {
                    MainDocumentPart mainPart = doc.MainDocumentPart;

                    IEnumerable<BookmarkStart> rs = from bm in mainPart.Document.Body.Descendants<BookmarkStart>() where bm.Name == bookMark select bm;

                    DocumentFormat.OpenXml.OpenXmlElement bookmark = rs.SingleOrDefault();

                    if (bookmark != null)
                    {
                        DocumentFormat.OpenXml.OpenXmlElement parent = bookmark.Parent;
                        AltChunk altchunk = new AltChunk();
                        string chunkId = bookMark + DateTime.Now.Minute + DateTime.Now.Second + DateTime.Now.Millisecond;
                        altchunk.Id = chunkId;
                        AlternativeFormatImportPart formatImport = doc.MainDocumentPart.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.Html, chunkId);
                        formatImport.FeedData(memoryStream);
                        parent.InsertBeforeSelf(altchunk);
                    }
                }
            }
            catch (Exception ext)
            {
                throw ext;
            }
        }
    }
}
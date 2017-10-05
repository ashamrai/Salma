using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Xml.Linq;
using WordToTFS;

namespace Salma2010
{
    #region Utils

    /// <summary>
    /// Utils
    /// </summary>
    internal sealed class Util
    {
        /// <summary>
        /// Stream to file
        /// </summary>
        /// <param name="inputStream"></param>
        /// <param name="outputFile"></param>
        /// <param name="fileMode"></param>
        /// <param name="sourceRectangle"></param>
        /// <returns></returns>
        public static string StreamToFile(Stream inputStream, string outputFile, FileMode fileMode, DocumentFormat.OpenXml.Drawing.SourceRectangle sourceRectangle)
        {
            try
            {
                if (inputStream == null)
                    throw new ArgumentNullException("inputStream");

                if (String.IsNullOrEmpty(outputFile))
                    throw new ArgumentException("Argument null or empty.", "outputFile");

                if (Path.GetExtension(outputFile).ToLower() == ".emf" || Path.GetExtension(outputFile).ToLower() == ".wmf")
                {
                    System.Drawing.Imaging.Metafile emf = new System.Drawing.Imaging.Metafile(inputStream);
                    System.Drawing.Rectangle cropRectangle;

                    double leftPercentage = (sourceRectangle != null && sourceRectangle.Left != null) ? ToPercentage(sourceRectangle.Left.Value) : 0;
                    double topPercentage = (sourceRectangle != null && sourceRectangle.Top != null) ? ToPercentage(sourceRectangle.Top.Value) : 0;
                    double rightPercentage = (sourceRectangle != null && sourceRectangle.Left != null) ? ToPercentage(sourceRectangle.Right.Value) : 0;
                    double bottomPercentage = (sourceRectangle != null && sourceRectangle.Left != null) ? ToPercentage(sourceRectangle.Bottom.Value) : 0;

                    cropRectangle = new System.Drawing.Rectangle(
                        (int)(emf.Width * leftPercentage),
                        (int)(emf.Height * topPercentage),
                        (int)(emf.Width - emf.Width * (leftPercentage + rightPercentage)),
                        (int)(emf.Height - emf.Height * (bottomPercentage + topPercentage)));

                    System.Drawing.Bitmap newBmp = new System.Drawing.Bitmap(cropRectangle.Width, cropRectangle.Height);
                    using (System.Drawing.Graphics graphic = System.Drawing.Graphics.FromImage(newBmp))
                    {
                        graphic.Clear(System.Drawing.Color.White);
                        graphic.DrawImage(emf, new System.Drawing.Rectangle(0, 0, cropRectangle.Width, cropRectangle.Height), cropRectangle, System.Drawing.GraphicsUnit.Pixel);
                    }
                    outputFile = outputFile.Replace(".emf", ".jpg");
                    newBmp.Save(outputFile, System.Drawing.Imaging.ImageFormat.Jpeg);
                    return outputFile;
                }
                else
                {
                    if (!File.Exists(outputFile))
                    {
                        using (FileStream outputStream = new FileStream(outputFile, fileMode, FileAccess.Write))
                        {
                            int cnt = 0;
                            const int LEN = 4096;
                            byte[] buffer = new byte[LEN];

                            while ((cnt = inputStream.Read(buffer, 0, LEN)) != 0)
                                outputStream.Write(buffer, 0, cnt);
                        }
                    }
                    return outputFile;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// To percentage
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static double ToPercentage(int value)
        {
            double result = 0;

            if (value.ToString().Length == 3)
                result = (double)value / 100 / 100;
            else if (value.ToString().Length == 5)
                result = (double)value / 1000 / 100;
            else
                result = (double)value / 1000 / 100;

            return result;
        }

        /// <summary>
        /// Get relative path
        /// </summary>
        /// <param name="imagePath"></param>
        /// <returns></returns>
        public static string GetRelativePath(string imagePath)
        {
            System.Uri uriImage = new Uri(imagePath);
            return uriImage.ToString();
        }

        /// <summary>
        /// Emu to pixels
        /// </summary>
        /// <param name="emu"></param>
        /// <returns></returns>
        public static int EmuToPixels(DocumentFormat.OpenXml.Int64Value emu)
        {
            if (emu != 0)
            {
                return (int)Math.Round((decimal)emu / 9525);
            }
            else
            {
                return 0;
            }
        }
    }

    #endregion

    /// http://hintdesk.com/toolc-docx-to-html-library/
    /// <summary>
    /// Docx To Html
    /// </summary>
    internal sealed partial class DocxToHtml
    {
        private static volatile DocxToHtml instance = null;
        private static object syncRoot = new Object();

        private WordprocessingDocument document = null;

        private string imageDirectory = "";

        /// <summary>
        /// Instance
        /// </summary>
        public static DocxToHtml Instance
        {
            get
            {
                if (instance == null)
                {
                    lock (syncRoot)
                    {
                        if (instance == null)
                            instance = new DocxToHtml();
                    }
                }

                return instance;
            }
        }

        /// <summary>
        /// Convert docx to html
        /// </summary>
        /// <param name="id"></param>
        /// <param name="fieldName"></param>
        public void ToHtmlExport(int id, string fieldName)
        {
            string folderName = Path.GetTempPath() + Guid.NewGuid().ToString();
            string fileName = folderName + "\\temp.docx";

            Globals.ThisAddIn.Application.ScreenUpdating = false;

            try
            {
                ExtendSelection();

                // create temp file
                Microsoft.Office.Interop.Word.Document doc = Globals.ThisAddIn.Application.Documents.Add(Visible: false, DocumentType:
                    Microsoft.Office.Interop.Word.WdDocumentType.wdTypeDocument);

                // get range
                doc.Range().FormattedText = Globals.ThisAddIn.Application.Selection.Range.FormattedText;
                PrepareRange(ref doc);

                // create temp folder
                if (Directory.Exists(folderName))
                    Directory.Delete(folderName, true);
                Directory.CreateDirectory(folderName);

                // save temp file
                doc.SaveAs(fileName, AddToRecentFiles: false);
                ((Microsoft.Office.Interop.Word._Document)doc).Close();

                // convert to html
                string html = WriteHtml(fileName, folderName);

                // add html
                TfsManager.Instance.ReplaceDetailsForWorkItem(id, fieldName, html);
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
            finally
            {
                Globals.ThisAddIn.Application.Selection.Collapse();
                Globals.ThisAddIn.Application.ScreenUpdating = true;

                try
                {
                    if (Directory.Exists(folderName))
                        Directory.Delete(folderName, true);
                }
                catch { }
            }
        }

        /// <summary>
        /// Convert docx to html
        /// </summary>
        /// <param name="id"></param>
        /// <param name="fieldName"></param>
        public void ToHtml(int id, string fieldName)
        {
            string folderName = Path.GetTempPath() + Guid.NewGuid().ToString();
            string fileName = folderName + "\\temp.docx";

            Globals.ThisAddIn.Application.ScreenUpdating = false;

            try
            {
                ExtendSelection();

                // create temp file
                Microsoft.Office.Interop.Word.Document doc = Globals.ThisAddIn.Application.Documents.Add(Visible: false, DocumentType:
                    Microsoft.Office.Interop.Word.WdDocumentType.wdTypeDocument);

                // get range
                doc.Range().FormattedText = Globals.ThisAddIn.Application.Selection.Range.FormattedText;
                PrepareRange(ref doc);

                // create temp folder
                if (Directory.Exists(folderName))
                    Directory.Delete(folderName, true);
                Directory.CreateDirectory(folderName);

                // save temp file
                doc.SaveAs(fileName, AddToRecentFiles: false);
                ((Microsoft.Office.Interop.Word._Document)doc).Close();

                // convert to html
                string html = WriteHtml(fileName, folderName);

                // add html
                TfsManager.Instance.ReplaceDetailsForWorkItem(id, fieldName, html);
                Globals.ThisAddIn.AddWIControl(id, fieldName);
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
            finally
            {
                
                Globals.ThisAddIn.Application.Selection.Collapse();
                Globals.ThisAddIn.Application.ScreenUpdating = true;

                try
                {
                    if (Directory.Exists(folderName))
                        Directory.Delete(folderName, true);
                }
                catch { }
            }
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
        /// Prepare range
        /// </summary>
        /// <param name="doc"></param>
        private static void PrepareRange(ref Microsoft.Office.Interop.Word.Document doc)
        {
            string r = Convert.ToChar(0x000D).ToString(); // \r

            for (int i = 1; i <= doc.ContentControls.Count; i++)
                doc.ContentControls[i].Delete(false);
            for (int i = 1; i <= doc.Comments.Count; i++)
                doc.Comments[i].DeleteRecursively();

            foreach (Microsoft.Office.Interop.Word.Paragraph paragraph in doc.Content.Paragraphs)
                if (paragraph.Range.Text.Trim() == string.Empty && paragraph.Range.InlineShapes.Count == 0)
                    paragraph.Range.Delete();
        }

        /// <summary>
        /// Write html
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="folderName"></param>
        /// <returns></returns>
        private string WriteHtml(string fileName, string folderName)
        {
            imageDirectory = folderName;
            string htmlBody = "";

            try
            {
                document = WordprocessingDocument.Open(fileName, false);

                foreach (OpenXmlElement element in document.MainDocumentPart.Document.Body.Elements<OpenXmlElement>())
                    if (element is Paragraph)
                        htmlBody += AddParagraph((Paragraph)element);
                    else if (element is Table)
                        htmlBody += AddTable((Table)element);

                document.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return GetHtml(htmlBody);
        }

        /// <summary>
        /// Add table
        /// </summary>
        /// <param name="table"></param>
        /// <returns></returns>
        private string AddTable(Table table)
        {
            try
            {
                string tableTemplate =
                    @"<table STYLE>
                        TABLE_CONTENT
                    </table>";
                    
                string htmlTableStyle = " width=\"100%\" border=\"1\" cellspacing=\"0\" cellpadding=\"0\" style=\"border-collapse:collapse;border:none;\"";
                string htmlRows = string.Empty;

                List<TableRow> rows = table.Elements<TableRow>().ToList();
                for (int indexRow = 0; indexRow < rows.Count(); indexRow++)
                {
                    TableRow row = rows[indexRow];
                    string htmlCells = "";
                    List<TableCell> cells = row.Elements<TableCell>().ToList();
                    for (int indexCell = 0; indexCell < cells.Count(); indexCell++)
                    {
                        TableCell cell = cells[indexCell];
                        string tableCellStyle = string.Format("<td width=\"{0}\" valign=\"top\" style=\"border:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt;\">", 100 / cells.Count);
                        htmlCells += tableCellStyle + AddParagraphs(cell.Descendants<Paragraph>()) + "</td>" + Environment.NewLine;
                    }
                    htmlRows += "<tr>" + htmlCells + "</tr>" + Environment.NewLine;
                }
                return tableTemplate.Replace("STYLE", htmlTableStyle).Replace("TABLE_CONTENT", htmlRows);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// Convert dxa to point
        /// </summary>
        /// <param name="dxa"></param>
        /// <returns></returns>
        private string ConvertDxaToPoint(string dxa)
        {
            return (Convert.ToDouble(dxa) / 20).ToString("f1").Replace(",", ".");
        }

        /// <summary>
        /// Add paragraphs
        /// </summary>
        /// <param name="paragraphs"></param>
        /// <returns></returns>
        private string AddParagraphs(IEnumerable<Paragraph> paragraphs)
        {
            string result = "";

            foreach (Paragraph paragraph in paragraphs)
                result += AddParagraph(paragraph);

            return result;
        }

        /// <summary>
        /// Add paragraph
        /// </summary>
        /// <param name="paragraph"></param>
        /// <returns></returns>
        private string AddParagraph(Paragraph paragraph)
        {
            try
            {
                string paragraphTemplate = "<p class=STYLE>PARAGRAPH_CONTENT</p>";
                string paragraphStyle = paragraph.ParagraphProperties != null ? AddFormatParagraph(paragraph.ParagraphProperties) : "\"\"";

                string paragraphContent = "";

                foreach (OpenXmlElement element in paragraph.Elements())
                    if (element is Run)
                        paragraphContent += AddRun((Run)element);
                    else if (element is Hyperlink)
                        paragraphContent += AddHyperlink((Hyperlink)element);

                return paragraphTemplate.Replace("STYLE", paragraphStyle)
                    .Replace("PARAGRAPH_CONTENT", paragraphContent == "" ? "&nbsp;" : paragraphContent) + Environment.NewLine;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// Add hyperlink
        /// </summary>
        /// <param name="hyperlink"></param>
        /// <returns></returns>
        private string AddHyperlink(Hyperlink hyperlink)
        {
            string styleTemplate = "<span style='STYLE_CONTENT'>HYPERLINK_CONTENT</span>";
            string style = "";
            List<string> cssStyleElements = new List<string>();

            HyperlinkRelationship relation = (from rel in document.MainDocumentPart.HyperlinkRelationships
                                              where rel.Id == hyperlink.Id
                                              select rel).FirstOrDefault() as HyperlinkRelationship;
            if (relation != null)
            {
                string result = "<a href=\"" + relation.Uri.ToString() + "\">" + hyperlink.Descendants<Run>().Select(x => x.InnerText).Aggregate((i, j) => i + j) + "</a>";

                FontSize fontSize = hyperlink.Descendants<FontSize>().FirstOrDefault();

                if (fontSize != null)
                    cssStyleElements.Add(string.Format("font-size:{0}pt", Convert.ToInt32(fontSize.Val.Value) / 2));

                foreach (string element in cssStyleElements)
                    style += element + " ";

                result = styleTemplate.Replace("HYPERLINK_CONTENT", result).Replace("STYLE_CONTENT", style);

                Italic italic = hyperlink.Descendants<Italic>().FirstOrDefault();

                if (italic != null)
                    result = AddFormat(result, "i");

                return result;
            }
            else
                return "";
        }

        /// <summary>
        /// Add run
        /// </summary>
        /// <param name="run"></param>
        /// <returns></returns>
        private string AddRun(Run run)
        {
            string runText = "";

            foreach (OpenXmlElement element in run.Elements())
            {
                if (element is TabChar)
                {
                    runText += "&emsp;&emsp;";
                }
                else if (element is Text)
                {
                    string textEncode = HttpUtility.HtmlEncode(element.InnerText);
                    if (textEncode.TakeWhile(x => x == ' ').Count() >= textEncode.Length / 2)
                    {
                        if (element.HasAttributes)
                        {
                            if (element.GetAttributes().Where(x => x.LocalName == "space" && x.Value == "preserve").FirstOrDefault() != null)
                                textEncode = textEncode.Replace(" ", "&nbsp;");
                        }
                    }
                    runText += textEncode;
                }
                else if (element is Break)
                {
                    runText += "<br/>";
                }
                else if (element is SymbolChar)
                {
                    runText += AddSymbolChar((SymbolChar)element);
                }
                else if (element is Picture)
                {
                    runText += AddPicture((Picture)element);
                }
                else if (element is Drawing)
                {
                    runText += AddDrawing((Drawing)element);
                }
                else if (element is AlternateContent)
                {
                    runText += AddAlternateContent((AlternateContent)element);
                }
                else if (element is RunProperties)
                {
                }
                else if (element is EmbeddedObject)
                {
                    runText += AddEmbeddedObject((EmbeddedObject)element);
                }
            }

            if (run.RunProperties != null)
                runText = AddFormat(runText, run.RunProperties);

            return runText;
        }

        /// <summary>
        /// Add embedded object
        /// </summary>
        /// <param name="embeddedObject"></param>
        /// <returns></returns>
        private string AddEmbeddedObject(EmbeddedObject embeddedObject)
        {
            foreach (OpenXmlElement element in embeddedObject.Elements())
                if (element is DocumentFormat.OpenXml.Vml.Shape)
                    return DrawShape((DocumentFormat.OpenXml.Vml.Shape)element);

            return "";
        }

        /// <summary>
        /// Add alternate content
        /// </summary>
        /// <param name="alternateContent"></param>
        /// <returns></returns>
        private string AddAlternateContent(AlternateContent alternateContent)
        {
            foreach (OpenXmlElement element in alternateContent.Elements())
                if (element is AlternateContentFallback)
                    return AddAlternateContentFallback((AlternateContentFallback)element);

            return "";
        }

        /// <summary>
        /// Add alternate content fallback
        /// </summary>
        /// <param name="alternateContentFallback"></param>
        /// <returns></returns>
        private string AddAlternateContentFallback(AlternateContentFallback alternateContentFallback)
        {
            foreach (OpenXmlElement element in alternateContentFallback.Elements())
                if (element is Picture)
                    return AddPicture((Picture)element);

            return "";
        }

        /// <summary>
        /// Add drawing
        /// </summary>
        /// <param name="drawing"></param>
        /// <returns></returns>
        private string AddDrawing(Drawing drawing)
        {
            foreach (OpenXmlElement element in drawing.Elements())
                if (element is Inline)
                    return AddInline((Inline)element);

            return "";
        }

        /// <summary>
        /// Add inline
        /// </summary>
        /// <param name="inline"></param>
        /// <returns></returns>
        private string AddInline(Inline inline)
        {
            string graphicName = "";
            int width = 0, height = 0;
            string fileName = "";

            foreach (OpenXmlElement element in inline.Elements())
            {
                if (element is DocProperties)
                    graphicName = ((DocProperties)element).Name.Value;
                else if (element is Extent)
                {
                    width = Util.EmuToPixels(((Extent)element).Cx);
                    height = Util.EmuToPixels(((Extent)element).Cy);
                }
                else if (element is DocumentFormat.OpenXml.Drawing.Graphic)
                {
                    fileName = AddGraphic((DocumentFormat.OpenXml.Drawing.Graphic)element);
                }
            }

            if (fileName != "")
                return string.Format("<img width=\"{0}\" height=\"{1}\" alt=\"{2}\" src=\"{3}\" />", width, height, graphicName, Util.GetRelativePath(fileName));
            else
                return "";
        }

        /// <summary>
        /// Add graphic
        /// </summary>
        /// <param name="graphic"></param>
        /// <returns></returns>
        private string AddGraphic(DocumentFormat.OpenXml.Drawing.Graphic graphic)
        {
            foreach (OpenXmlElement element in graphic.Elements())
                if (element is DocumentFormat.OpenXml.Drawing.GraphicData)
                    return AddGraphicData((DocumentFormat.OpenXml.Drawing.GraphicData)element);

            return "";
        }

        /// <summary>
        /// Add graphic data
        /// </summary>
        /// <param name="graphicData"></param>
        /// <returns></returns>
        private string AddGraphicData(DocumentFormat.OpenXml.Drawing.GraphicData graphicData)
        {
            foreach (OpenXmlElement element in graphicData.Elements())
                if (element is DocumentFormat.OpenXml.Drawing.Pictures.Picture)
                    return AddPicture((DocumentFormat.OpenXml.Drawing.Pictures.Picture)element);

            return "";
        }

        /// <summary>
        /// Add picture
        /// </summary>
        /// <param name="picture"></param>
        /// <returns></returns>
        private string AddPicture(DocumentFormat.OpenXml.Drawing.Pictures.Picture picture)
        {
            foreach (OpenXmlElement element in picture.Elements())
            {
                if (element is DocumentFormat.OpenXml.Drawing.Pictures.BlipFill)
                {
                    DocumentFormat.OpenXml.Drawing.Blip blip = ((DocumentFormat.OpenXml.Drawing.Pictures.BlipFill)element).Blip;

                    if (blip != null)
                    {
                        OpenXmlPart image = document.MainDocumentPart.GetPartById(blip.Embed.Value);
                        string fileName = Path.Combine(imageDirectory, Path.GetFileName(image.Uri.ToString()));
                        fileName = Util.StreamToFile(image.GetStream(), fileName, FileMode.CreateNew, null);
                        return fileName;
                    }
                }
            }

            return "";
        }

        /// <summary>
        /// Add symbol char
        /// </summary>
        /// <param name="symbolChar"></param>
        /// <returns></returns>
        private string AddSymbolChar(SymbolChar symbolChar)
        {
            string result = "";

            switch (symbolChar.Char.Value)
            {
                case "F0E7":
                    result = ((char)8592).ToString();
                    break;

                case "F0E0":
                    result = ((char)8594).ToString();
                    break;

                default:
                    break;
            }

            return result;
        }

        /// <summary>
        /// Add picture
        /// </summary>
        /// <param name="picture"></param>
        /// <returns></returns>
        private string AddPicture(Picture picture)
        {
            foreach (OpenXmlElement element in picture.Elements())
                if (element is DocumentFormat.OpenXml.Vml.RoundRectangle)
                    return DrawRoundRectangle((DocumentFormat.OpenXml.Vml.RoundRectangle)element);
                else if (element is DocumentFormat.OpenXml.Vml.Group)
                    return DrawGroup((DocumentFormat.OpenXml.Vml.Group)element);
                else if (element is DocumentFormat.OpenXml.Vml.Shape)
                    return DrawShape((DocumentFormat.OpenXml.Vml.Shape)element);

            return "";
        }

        /// <summary>
        /// Draw group
        /// </summary>
        /// <param name="group"></param>
        /// <returns></returns>
        private string DrawGroup(DocumentFormat.OpenXml.Vml.Group group)
        {
            string result = "";

            foreach (OpenXmlElement element in group.Elements())
                if (element is DocumentFormat.OpenXml.Vml.Shape)
                    result += DrawShape((DocumentFormat.OpenXml.Vml.Shape)element);
                else if (element is DocumentFormat.OpenXml.Vml.Line)
                    result += DrawLine((DocumentFormat.OpenXml.Vml.Line)element);

            return result;
        }

        /// <summary>
        /// Draw shape
        /// </summary>
        /// <param name="shape"></param>
        /// <returns></returns>
        private string DrawShape(DocumentFormat.OpenXml.Vml.Shape shape)
        {
            string style = shape.GetAttributes().Where(x => x.LocalName == "style").FirstOrDefault().Value;
            string position = GetValueOfProperty("position", style);

            string styleLeft = GetValueOfProperty("left", style);
            int marginLeft = ConvertToPixel(GetValueOfProperty("margin-left", style));
            int marginTop = ConvertToPixel(GetValueOfProperty("margin-top", style));
            int width = ConvertToPixel(GetValueOfProperty("width", style));
            int height = ConvertToPixel(GetValueOfProperty("height", style));
            string graphicName = shape.Id;

            foreach (OpenXmlElement element in shape.Elements())
                if (element is DocumentFormat.OpenXml.Vml.ImageData)
                    return DrawImageData(position, marginLeft, marginTop, width, height, (DocumentFormat.OpenXml.Vml.ImageData)element);
                else if (element is DocumentFormat.OpenXml.Vml.TextBox)
                    return AddTextBox((DocumentFormat.OpenXml.Vml.TextBox)element);

            return "";
        }

        /// <summary>
        /// Add text box
        /// </summary>
        /// <param name="textBox"></param>
        /// <returns></returns>
        private string AddTextBox(DocumentFormat.OpenXml.Vml.TextBox textBox)
        {
            foreach (OpenXmlElement element in textBox.Elements())
                if (element is TextBoxContent)
                    return AddTextBoxContent((TextBoxContent)element);

            return "";
        }

        /// <summary>
        /// Add text box content
        /// </summary>
        /// <param name="textBoxContent"></param>
        /// <returns></returns>
        private string AddTextBoxContent(TextBoxContent textBoxContent)
        {
            string result = "";

            foreach (OpenXmlElement element in textBoxContent.Elements())
                if (element is Paragraph)
                    result += AddParagraph((Paragraph)element);
 
            return result;
        }

        /// <summary>
        /// Draw image data
        /// </summary>
        /// <param name="position"></param>
        /// <param name="marginLeft"></param>
        /// <param name="marginTop"></param>
        /// <param name="width"></param>
        /// <param name="height"></param>
        /// <param name="imageData"></param>
        /// <returns></returns>
        private string DrawImageData(string position, int marginLeft, int marginTop, int width, int height, DocumentFormat.OpenXml.Vml.ImageData imageData)
        {
            OpenXmlPart image = document.MainDocumentPart.GetPartById(imageData.RelationshipId);
            string fileName = Path.Combine(imageDirectory, Path.GetFileName(image.Uri.ToString()));
            fileName = Util.StreamToFile(image.GetStream(), fileName, FileMode.CreateNew, null);
            return string.Format("<span style='position:{0};margin-left:{1}px;margin-top:{2}px;width:{3}px;height:{4}px'><img width=\"{3}\" height=\"{4}\" alt=\"{5}\" src=\"{6}\"/></span>", 
                position, marginLeft, marginTop, width, height, Path.GetFileName(fileName), Util.GetRelativePath(fileName));
        }

        /// <summary>
        /// Convert to pixel
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        private int ConvertToPixel(string value)
        {
            if (value != null)
            {
                if (value.Contains("pt"))
                    return ConvertPointToPixel(value);
                else if (value.Contains("in"))
                    return ConvertInchToPixel(value);
                else
                    return ConvertDxaToPixel(value);
            }
            else
                return 0;
        }

        /// <summary>
        /// Convert inch to pixel
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        private int ConvertInchToPixel(string value)
        {
            value = value.Replace("in", "");
            return ConvertPointToPixel((Convert.ToDouble(value) * 72).ToString());
        }

        /// <summary>
        /// Convert dxa to pixel
        /// </summary>
        /// <param name="dxa"></param>
        /// <returns></returns>
        private int ConvertDxaToPixel(string dxa)
        {
            if (dxa != null)
                return ((int)(Convert.ToDouble(dxa) / 20) * 96 / 72);
            else
                return 0;
        }

        /// <summary>
        /// Draw round rectangle
        /// </summary>
        /// <param name="roundRectangle"></param>
        /// <returns></returns>
        private string DrawRoundRectangle(DocumentFormat.OpenXml.Vml.RoundRectangle roundRectangle)
        {
            string fileName = Path.Combine(imageDirectory, roundRectangle.GetAttributes().Where(x => x.LocalName == "id").FirstOrDefault().Value + ".jpg");
            string style = roundRectangle.GetAttributes().Where(x => x.LocalName == "style").FirstOrDefault().Value;
            string position = GetValueOfProperty("position", style);
            int marginLeft = ConvertPointToPixel(GetValueOfProperty("margin-left", style));
            int marginTop = ConvertPointToPixel(GetValueOfProperty("margin-top", style));
            int width = ConvertPointToPixel(GetValueOfProperty("width", style));
            int height = ConvertPointToPixel(GetValueOfProperty("height", style));

            System.Drawing.Bitmap newBmp = new System.Drawing.Bitmap(width, height);

            using (System.Drawing.Graphics graphic = System.Drawing.Graphics.FromImage(newBmp))
            {
                System.Drawing.Pen pen = new System.Drawing.Pen(ConvertToColor(roundRectangle.StrokeColor.Value), ConvertPointToPixel(roundRectangle.StrokeWeight.Value));
                pen.Alignment = System.Drawing.Drawing2D.PenAlignment.Inset;
                System.Drawing.SolidBrush solidBrush = new System.Drawing.SolidBrush(ConvertToColor(roundRectangle.FillColor.Value));
                System.Drawing.Rectangle rectangle = new System.Drawing.Rectangle(0, 0, width, height);
                graphic.FillRectangle(solidBrush, rectangle);
                graphic.DrawRectangle(pen, rectangle);
            }
            newBmp.Save(fileName, System.Drawing.Imaging.ImageFormat.Jpeg);

            return string.Format("<span style='position:{0};margin-left:{1}px;margin-top:{2}px;width:{3}px;height:{4}px'><img width=\"{3}\" height=\"{4}\" alt=\"{5}\" src=\"{6}\"/>", 
                position, marginLeft, marginTop, width, height, Path.GetFileName(fileName), Util.GetRelativePath(fileName));
        }

        /// <summary>
        /// Get value of property
        /// </summary>
        /// <param name="name"></param>
        /// <param name="style"></param>
        /// <returns></returns>
        private string GetValueOfProperty(string name, string style)
        {
            string value = style.Split(";".ToCharArray()).Where(x => x.Contains(name)).FirstOrDefault();
            if (value != null)
                return value.Split(":".ToCharArray())[1];
            else
                return null;
        }

        /// <summary>
        /// Draw line
        /// </summary>
        /// <param name="line"></param>
        /// <returns></returns>
        private string DrawLine(DocumentFormat.OpenXml.Vml.Line line)
        {
            string fileName = Path.Combine(imageDirectory, line.GetAttributes().Where(x => x.LocalName == "id").FirstOrDefault().Value + ".jpg");
            string style = line.GetAttributes().Where(x => x.LocalName == "style").FirstOrDefault().Value;
            string position = GetValueOfProperty("position", style);
            int marginLeft = ConvertPointToPixel(GetValueOfProperty("margin-left", style));
            int marginTop = ConvertPointToPixel(GetValueOfProperty("margin-top", style));
            int width = (int)(ConvertToPixel(line.To.Value.Split(",".ToCharArray())[0]) - ConvertToPixel(line.From.Value.Split(",".ToCharArray())[0]));
            int height = (int)(ConvertToPixel(line.To.Value.Split(",".ToCharArray())[1]) - ConvertToPixel(line.From.Value.Split(",".ToCharArray())[1]));
            string strokeWeight = line.StrokeWeight != null ? line.StrokeWeight.Value : "3pt";

            if (height == 0)
                height = ConvertPointToPixel(strokeWeight);

            System.Drawing.Bitmap newBmp = new System.Drawing.Bitmap(width, height);

            using (System.Drawing.Graphics graphic = System.Drawing.Graphics.FromImage(newBmp))
            {
                System.Drawing.Pen pen = new System.Drawing.Pen(ConvertToColor(line.StrokeColor != null ? line.StrokeColor.Value : "black"), height);
                graphic.DrawLine(pen, 0, 0, newBmp.Width, newBmp.Height);
            }

            newBmp.Save(fileName, System.Drawing.Imaging.ImageFormat.Jpeg);

            return string.Format("<span style='position:{0};margin-left:{1}px;margin-top:{2}px;width:{3}px;height:{4}px'><img width=\"{3}\" height=\"{4}\" alt=\"{5}\" src=\"{6}\"/></span>", 
                position, marginLeft, marginTop, width, height, Path.GetFileName(fileName), Util.GetRelativePath(fileName)); 
        }

        /// <summary>
        /// Convert point to pixel
        /// </summary>
        /// <param name="point"></param>
        /// <returns></returns>
        private int ConvertPointToPixel(string point)
        {
            if (point != null)
                return (int)(double.Parse(point.Replace("pt", ""), CultureInfo.InvariantCulture) * 96 / 72);
            else
                return 0;
        }

        /// <summary>
        /// Convert to color
        /// </summary>
        /// <param name="strokeColor"></param>
        /// <returns></returns>
        private System.Drawing.Color ConvertToColor(string strokeColor)
        {
            if (strokeColor.Contains("#"))
            {
                strokeColor = strokeColor.Split(" ".ToCharArray())[0];
                return System.Drawing.ColorTranslator.FromHtml(strokeColor);
            }
            else
                return System.Drawing.Color.FromName(strokeColor.Split(" ".ToCharArray())[0]);
        }

        /// <summary>
        /// Add format paragraph
        /// </summary>
        /// <param name="paragraphProperties"></param>
        /// <returns></returns>
        private string AddFormatParagraph(ParagraphProperties paragraphProperties)
        {
            List<string> cssStyleElements = new List<string>();

            string paragraphStyle = "";
         
            if (paragraphProperties != null)
            {
                ParagraphStyleId paragraphStyleId = paragraphProperties.Descendants<ParagraphStyleId>().FirstOrDefault();
                
                if (paragraphStyleId != null)
                    paragraphStyle = paragraphStyleId.Val;
            }

            paragraphStyle = "\"" + paragraphStyle + "\"";

           
            Indentation indentation = paragraphProperties.Indentation;

            /*
            if (indentation != null)
            {
                if (indentation.Left != null)
                    cssStyleElements.Add(string.Format("margin-left:{0}pt", ConvertDxaToPoint(indentation.Left.Value)));
                else if (indentation.FirstLine != null)
                    //cssStyleElements.Add(string.Format("text-indent:{0}pt", ConvertDxaToPoint(indentation.FirstLine.Value)));
                    cssStyleElements.Add(string.Format("margin-left:{0}pt", ConvertDxaToPoint(indentation.FirstLine.Value)));
            }*/

            NumberingProperties numberingProperties = paragraphProperties.NumberingProperties;

            if (numberingProperties != null && numberingProperties.NumberingLevelReference != null && numberingProperties.NumberingId != null)
            {
                int level = numberingProperties.NumberingLevelReference.Val.Value;
                int numId = numberingProperties.NumberingId.Val.Value;

                int abstractNumId = document.MainDocumentPart.NumberingDefinitionsPart.Numbering.Descendants<NumberingInstance>().Where(x => x.NumberID == numId).FirstOrDefault().AbstractNumId.Val.Value;
                indentation = document.MainDocumentPart.NumberingDefinitionsPart.Numbering.Descendants<AbstractNum>().Where(x => x.AbstractNumberId == abstractNumId).FirstOrDefault().Descendants<Level>().Where(x => x.LevelIndex.Value == level).FirstOrDefault().Descendants<Indentation>().FirstOrDefault();
               
                Level numberingLevel = document.MainDocumentPart.NumberingDefinitionsPart.Numbering.Descendants<AbstractNum>().Where(x => x.AbstractNumberId == abstractNumId).FirstOrDefault().Descendants<Level>().Where(x => x.LevelIndex.Value == level).FirstOrDefault();
               
                if (numberingLevel.LevelText.Val.Value.Contains("%"))
                    cssStyleElements.Add(string.Format("mso-list:l{0} level{1} ol", numId, level + 1));
                else
                    cssStyleElements.Add(string.Format("mso-list:l{0} level{1} ul", numId, level + 1));
            }

            string result = "";

            foreach (string element in cssStyleElements)
                result += element + ";";

            if (result != "")
                result = paragraphStyle + string.Format(" style='{0}'", result);
            else
                result = paragraphStyle;

            return result.Trim();
        }

        /// <summary>
        /// Add format
        /// </summary>
        /// <param name="runText"></param>
        /// <param name="runProperties"></param>
        /// <returns></returns>
        private string AddFormat(string runText, RunProperties runProperties)
        {
            string spanStyle = "";

            if (runProperties.RunStyle != null)
                spanStyle = "<span class=" + runProperties.RunStyle.Val.Value + " style='SPAN_STYLE'>{0}</span>";
            else
                spanStyle = "<span style='SPAN_STYLE'>{0}</span>";

            List<string> cssStyleElements = new List<string>();

            if (runProperties.Bold != null)
                cssStyleElements.Add("font-weight: bold");
            if (runProperties.Underline != null)
                cssStyleElements.Add("text-decoration: underline");
            if (runProperties.Italic != null)
                cssStyleElements.Add("font-style: italic");

            if (cssStyleElements.Count > 0)
            {
                string spanStyleElement = "";
                foreach (string element in cssStyleElements)
                    spanStyleElement += element + ";";
                spanStyle = spanStyle.Replace("SPAN_STYLE", spanStyleElement);
            }
            else
                spanStyle = spanStyle.Replace("SPAN_STYLE", "");

            runText = string.Format(spanStyle, runText);

            return runText;
        }

        /// <summary>
        /// Add format
        /// </summary>
        /// <param name="runText"></param>
        /// <param name="tag"></param>
        /// <returns></returns>
        private string AddFormat(string runText, string tag)
        {
            return string.Format("<{0}>{1}</{0}>", tag, runText);
        }
    }
}
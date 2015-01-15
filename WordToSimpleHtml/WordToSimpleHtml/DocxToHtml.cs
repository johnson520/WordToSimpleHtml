﻿using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;

namespace WordToSimpleHtml
{
    /// <summary>
    /// Summary description for DocxToHtml
    /// </summary>
    public class DocxToHtml
    {
        private static readonly Regex rxRelationship =
			new Regex(@"<Relationship\s+Id=""(?<key>[^""]+)"".+?Type=""(?<type>[^""]+)"".+?Target=""(?<value>[^""]+)""(?:\s+TargetMode=""(?<mode>[^""]+)"")?.*?/>",
                      RegexOptions.Singleline | RegexOptions.Compiled);

        private static readonly Regex rxBody = new Regex(@"<w:body\b[^>]*>(?<inner>.+?)</w:body>", RegexOptions.Singleline | RegexOptions.Compiled);
        private static readonly Regex rxParagraph = new Regex(@"<w:p\b[^>]*>(?<inner>.+?)</w:p>", RegexOptions.Singleline | RegexOptions.Compiled);
        private static readonly Regex rxText = new Regex(@"<w:(?<br>br)/>|<w:(?<tab>tab)/>|<w:t\b[^>]*>(?<inner>.+?)</w:t>", RegexOptions.Singleline | RegexOptions.Compiled);
        private static readonly Regex rxRun = new Regex(@"<(?<tag>w:(?:fake)?r)\b[^>]*>(?<inner>.+?)</\k<tag>>", RegexOptions.Singleline | RegexOptions.Compiled);
        private static readonly Regex rxRunProp = new Regex(@"<w:rPr\b[^>]*>(?<inner>.+?)</w:rPr>", RegexOptions.Singleline | RegexOptions.Compiled);

        private static readonly Regex rxPStyle = new Regex(@"<w:pStyle\s+w:val=""(?<style>[^""]+)""\s*/>", RegexOptions.Singleline | RegexOptions.Compiled);

/*
		private static readonly Regex rxNumberList = new Regex(@"<w:pPr>\s*<w:pStyle\s*w:val=""ListParagraph""/>\s*<w:numPr>", RegexOptions.Singleline | RegexOptions.Compiled);
*/

        private static readonly Regex rxHeadingStyle = new Regex(@"^Heading(?<n>\d+)$", RegexOptions.Singleline | RegexOptions.Compiled);

        private static readonly Regex rxHyperlink = new Regex(@"<w:hyperlink\s+r:id=""(?<relid>[^""]+)""(?:\s+w:anchor=""(?<anchor>[^""]+)"")?[^>]*>(?<inner>.+?)</w:hyperlink>",
                                                              RegexOptions.Singleline | RegexOptions.Compiled);

        private static readonly Regex rxDrawingImage = new Regex(@"<w:drawing>(?<inner>.+?)</w:drawing>",
                                                                 RegexOptions.Singleline | RegexOptions.Compiled);

        private static readonly Regex rxTable = new Regex(@"<w:tbl>(?<inner>.+?)</w:tbl>", RegexOptions.Singleline | RegexOptions.Compiled);
        private static readonly Regex rxTableRow = new Regex(@"<w:tr\b[^>]*>(?<inner>.+?)</w:tr>", RegexOptions.Singleline | RegexOptions.Compiled);
        private static readonly Regex rxTableCell = new Regex(@"<w:tc>(?<inner>.+?)</w:tc>", RegexOptions.Singleline | RegexOptions.Compiled);

        private static readonly Regex rxGridSpan = new Regex(@"<w:gridSpan\s+w:val=""(?<colspan>\d+)""/>");

        private readonly Dictionary<string, string> rels = new Dictionary<string, string>();

        private string htmlDir;

	    public delegate void ErrorLogger(string s);

		private static void noLogger(string s)
		{
		}

	    private readonly ErrorLogger logError = noLogger;

		public DocxToHtml(ErrorLogger logger)
		{
			logError = logger;
		}

		public void Convert(string docxFile, string htmlFile, string imageFilePrefix, out string foundTitle)
		{
		    foundTitle = null;

			if (!File.Exists(docxFile))
			{
				logError(string.Format("Word file '{0}' does not exist.", docxFile));
				return;
			}

			htmlDir = Path.GetDirectoryName(htmlFile);
			if (htmlDir == null || !Directory.Exists(htmlDir))
			{
				logError(string.Format("HTML file directory '{0}' does not exist", htmlDir ?? "null"));
			}

			htmlDir += "\\";

			try
			{
				string content;
				using (var p = Package.Open(docxFile))
				{
					LoadRels(p, imageFilePrefix);
					content = ReadAllPart(p, "/word/document.xml");
					p.Close();
				}

				var html = ConvertContentToHtml(content, out foundTitle);

			    if (!string.IsNullOrEmpty(foundTitle))
			        html = string.Format("<h1 class='title'>{0}</h1>" + Environment.NewLine, foundTitle) + html;

				File.WriteAllText(htmlFile, html, Encoding.UTF8);
			}
			catch (Exception theE)
			{
				File.WriteAllText(htmlFile, "Exception occurred converting docx to html: " + theE.Message + " at " + theE.StackTrace);
				throw;
			}
		}

        private void LoadRels(Package p, string imageFilePrefix)
        {
            var relsXml = ReadAllPart(p, "/word/_rels/document.xml.rels");

            for (var m = rxRelationship.Match(relsXml); m.Success; m = m.NextMatch())
            {
                var relKey = m.Groups["key"].Value;
                var relValue = m.Groups["value"].Value;

                if (m.Groups["type"].Value.EndsWith("/image"))
                {
                    var imageFileName = ImageFileName(imageFilePrefix, relValue);

                    try
                    {
                        SaveImage(p, relValue, m.Groups["mode"].Value == "External", htmlDir + imageFileName);
                        rels.Add(relKey, imageFileName);
                    }
                    catch (Exception exc)
                    {
						logError(exc.Message);
                        rels.Add(relKey, relValue);
                    }
                }
                else
                {
                    rels.Add(relKey, relValue);
                }
            }
        }

        private static void SaveImage(Package p, string relValue, bool imageIsExternal, string imageFullPath)
        {
            if (imageIsExternal)
            {
                using (var wc = new WebClient())
                {
                    wc.DownloadFile(relValue, imageFullPath);
                }
            }
            else
            {
                var pp = p.GetPart(new Uri("/word/" + relValue, UriKind.Relative));
                using (var pps = pp.GetStream())
                {
                    using (var br = new BinaryReader(pps))
                        File.WriteAllBytes(imageFullPath, br.ReadBytes(System.Convert.ToInt32(pps.Length)));

                    pps.Close();
                }
            }
        }

        private string ImageFileName(string imageFilePrefix, string relValue)
        {
            var name = imageFilePrefix + Path.GetFileName(relValue);

            var uniqueSuffix = 0;
            while (File.Exists(htmlDir + name))
                name = string.Format("{0}{1}-{2}{3}", imageFilePrefix, Path.GetFileNameWithoutExtension(relValue), ++uniqueSuffix, Path.GetExtension(relValue));

            return name;
        }

        private string ConvertContentToHtml(string content, out string foundTitle)
        {
            foundTitle = null;

            var bodyMatch = rxBody.Match(content);
            if (!bodyMatch.Success)
                return string.Empty;

            var bodyContent = bodyMatch.Groups["inner"].Value;

            bodyContent = ReplaceImages(bodyContent);
            bodyContent = ReplaceHyperlinks(bodyContent);
            bodyContent = ReplaceTables(bodyContent);

            bodyContent = CollectParagraphs(bodyContent);

            return FinalCleanup(bodyContent, out foundTitle);
        }

        private static readonly Regex rxTableCleanup = new Regex(@"<p>\s*<table>", RegexOptions.Compiled | RegexOptions.Singleline);
        private static readonly Regex rxTableEndCleanup = new Regex(@"</table>\s*</p>", RegexOptions.Compiled | RegexOptions.Singleline);

		private static readonly Regex rxInnerText = new Regex(@">(?<inner>[^>]*)<", RegexOptions.Compiled | RegexOptions.Singleline);
		private static readonly Regex rxWebsomething = new Regex(@"\b[wW]eb(?<something>(?:site|page)s?)\b", RegexOptions.Compiled | RegexOptions.Singleline);
		private static readonly Regex rxWeb = new Regex(@"\bweb\b", RegexOptions.Compiled | RegexOptions.Singleline);
		private static readonly Regex rxInternet = new Regex(@"\binternet\b", RegexOptions.Compiled | RegexOptions.Singleline);
        private static readonly Regex rxQuotePunctuation = new Regex(@"”(?<punc>[,\.])", RegexOptions.Compiled | RegexOptions.Singleline);
        private static readonly Regex rxTestDrive = new Regex(@"\btestdrive\b", RegexOptions.IgnoreCase | RegexOptions.Compiled | RegexOptions.Singleline);

        private static readonly Regex rxLonelyPsInTds = new Regex(@"<td>\s*<p>(?<inner>(?:(?!<p>).)*?)</p>\s*</td>", RegexOptions.Compiled | RegexOptions.Singleline);

        private static readonly Regex rxEmptyPs = new Regex(@"<p[^>]*>\s*</p>\s*", RegexOptions.Compiled | RegexOptions.Singleline);

        private static readonly Regex rxTitleP = new Regex(@"^\s*<p\s+class=""Title"">\s*(?<inner>.*?)\s*</p>\s*", RegexOptions.Compiled | RegexOptions.Singleline);
        private static readonly Regex rxInitialH1 = new Regex(@"^\s*<h1[^>]*>\s*(?<inner>.*?)\s*</h1>\s*", RegexOptions.Compiled | RegexOptions.Singleline);

        private static readonly Regex rxImgInP = new Regex(@"<p>\s*(?:<[bi]>)*\s*<img\s+src=""(?<src>[^""]+)""\s*/>\s*(?:<br\s?/>)*\s*(?<caption>.*?)\s*(?:</[bi]>\s*)*</p>",
            RegexOptions.Compiled | RegexOptions.Singleline);
        private static readonly Regex rxNoBI = new Regex("</?[bi]>", RegexOptions.Compiled);
        private static readonly Regex rxNoTags = new Regex("<[^>]*>", RegexOptions.Compiled);

        private static readonly Regex rxUnneededBr = new Regex(@"(?<keep><p>)(?:\s*<br\s*/>)+|(?:<br\s*/>\s*)+(?<keep></p>)");

        private static readonly Regex rxListAfterP = new Regex(@"<p>(?<inner>(?:(?!:?</p>).)*):</p>\s*<ul>");

        private static string FinalCleanup(string bodyContent, out string foundTitle)
        {
            bodyContent = rxTableCleanup.Replace(bodyContent, "<table border=\"1\" cellpadding=\"4\" cellspacing=\"0\" style=\"border-collapse: collapse; margin-left: auto; margin-right: auto;\">");
            bodyContent = rxTableEndCleanup.Replace(bodyContent, "</table>");

            bodyContent = rxLonelyPsInTds.Replace(bodyContent, "<td>${inner}</td>");

            bodyContent = rxImgInP.Replace(bodyContent,
                                           m =>
                                           string.Format(
                                               "<p style=\"text-align: center; font-size: 8.25pt; font-style: italic;\"><img style=\"max-width: 95%; border: 1px solid #999; box-shadow: 3px 3px 8px 0px rgba(0, 0, 0, 0.25);\" alt=\"{2}\" src=\"{1}\" /><br />{0}</p>",
                                               rxNoBI.Replace(m.Groups["caption"].Value, string.Empty), m.Groups["src"].Value, rxNoTags.Replace(m.Groups["caption"].Value, string.Empty)));

            bodyContent = rxUnneededBr.Replace(bodyContent, "${keep}");
            bodyContent = rxEmptyPs.Replace(bodyContent, string.Empty);
            
            //  there are two ways we might find a title in the word document: an initial title paragraph or an initial h1
            var titleMatch = rxTitleP.Match(bodyContent);
            if (titleMatch.Success)
            {
                foundTitle = titleMatch.Groups["inner"].Value;
                bodyContent = rxTitleP.Replace(bodyContent, string.Empty);
            }
            else
            {
                foundTitle = null;
            }

            var h1Match = rxInitialH1.Match(bodyContent);
            if (h1Match.Success)
            {
                foundTitle = h1Match.Groups["inner"].Value;
                bodyContent = rxInitialH1.Replace(bodyContent, string.Empty);
            }

            bodyContent = rxListAfterP.Replace(bodyContent, "<p style=\"margin-bottom:0;\">${inner}:</p>" + Environment.NewLine + "<ul style=\"margin-top:0;\">");

			bodyContent = rxInnerText.Replace(bodyContent, InnerTextCleanup);

            return bodyContent;
        }

		private static string InnerTextCleanup(Match m)
		{
			var s = m.Groups["inner"].Value;

			s = rxWebsomething.Replace(s, "Web ${something}");
			s = rxWeb.Replace(s, "Web");
			s = rxInternet.Replace(s, "Internet");
            s = rxTestDrive.Replace(s, "Test Drive");
            s = rxQuotePunctuation.Replace(s, "${punc}”");

			return ">" + s + "<";
		}

        private static string CollectParagraphs(string bodyContent)
        {
            var sb = new StringBuilder();
            AppendParagraphs(sb, bodyContent);

            return sb.ToString();
        }

        private static void AppendParagraphs(StringBuilder sb, string bodyContent)
        {
            var inList = false;
            string listTag = null;

            for (var pMatch = rxParagraph.Match(bodyContent); pMatch.Success; pMatch = pMatch.NextMatch())
            {
                string tag = "p", style = null;

                var pInner = pMatch.Groups["inner"].Value;

                if (string.IsNullOrEmpty(pInner))
                    continue;

                var pStyleMatch = rxPStyle.Match(pInner);
                if (pStyleMatch.Success)
                {
                    style = pStyleMatch.Groups["style"].Value;

                    if (style.Contains("Normal"))
                    {
                        style = string.Empty;
                    }
                    else if (style == "ListParagraph")
                    {
                        if (!inList)
                        {
                            //	to figure out number vs. bullet, we have to redirect into numbering.xml and lookup the <w:numId w:val="1"/>.
                            //	<w:ilvl w:val="0"/> could tell us nestedness.
                            //listTag = rxNumberList.IsMatch(pInner) ? "ol" : "ul";
                            listTag = "ul";
                            sb.Append("<" + listTag + ">" + Environment.NewLine);
                            inList = true;
                        }

                        tag = "li";
                        style = null;
                    }
                    else
                    {
                        var headingMatch = rxHeadingStyle.Match(style);
                        if (headingMatch.Success)
                        {
                            tag = "h" + headingMatch.Groups["n"].Value;
                            style = null;
                        }
                    }
                }

                if (inList && tag != "li")
                {
                    sb.Append("</" + listTag + ">" + Environment.NewLine);
                    inList = false;
                }

                //  do some work to avoid outputting empty elements (empty <ul></ul> will still be output)
                var preTagLocation = sb.Length;

                sb.AppendFormat("<{0}{1}>", tag, string.IsNullOrEmpty(style) ? string.Empty : string.Format(" class=\"{0}\"", style));

                var preTextLocation = sb.Length;

                AppendInnerText(sb, pInner);

                if (sb.Length == preTextLocation)
                    sb.Remove(preTagLocation, sb.Length - preTagLocation);
                else
                    sb.AppendFormat("</{0}>" + Environment.NewLine, tag);
            }
        }

        private static void AppendInnerText(StringBuilder sb, string content)
        {
            var inBoldTag = 0;
            var inItalicTag = 0;

            for (var runMatch = rxRun.Match(content); runMatch.Success; runMatch = runMatch.NextMatch())
            {
                var runBold = false;
                var runItalic = false;

                var runPropMatch = rxRunProp.Match(runMatch.Groups["inner"].Value);
                if (runPropMatch.Success)
                {
                    runBold = runPropMatch.Groups["inner"].Value.Contains("<w:b/>");
                    runItalic = runPropMatch.Groups["inner"].Value.Contains("<w:i/>");
                }

                //  handle the case where we're in bold and italic and we're leaving it. get the order of end tags correct
                if (inBoldTag > 0 && inItalicTag > 0 && !(runBold && runItalic))
                {
                    sb.Append(inItalicTag > inBoldTag ? "</i></b>" : "</b></i>");
                    inItalicTag = 0;
                    inBoldTag = 0;
                }

                if (inItalicTag > 0 && !runItalic)
                {
                    sb.Append("</i>");
                    inItalicTag = 0;
                }

                if (inBoldTag > 0 && !runBold)
                {
                    sb.Append("</b>");
                    inBoldTag = 0;
                }

                if (runBold && inBoldTag == 0)
                {
                    sb.AppendFormat("<b>");
                    inBoldTag = sb.Length;
                }

                if (runItalic && inItalicTag == 0)
                {
                    sb.AppendFormat("<i>");
                    inItalicTag = sb.Length;
                }

                for (var tMatch = rxText.Match(runMatch.Groups["inner"].Value); tMatch.Success; tMatch = tMatch.NextMatch())
                {
                    sb.Append(tMatch.Groups["br"].Value == "br" ? "<br/>" : tMatch.Groups["tab"].Value == "tab" ? "&nbsp;&nbsp;&nbsp;&nbsp;" : tMatch.Groups["inner"].Value);
                }
            }

            if (inItalicTag > 0 && inBoldTag > 0)
                sb.Append(inItalicTag > inBoldTag ? "</i></b>" : "</b></i>");
            else if (inItalicTag > 0)
                sb.Append("</i>");
            else if (inBoldTag > 0)
                sb.Append("</b>");
        }

        private string ReplaceHyperlinks(string content)
        {
            return rxHyperlink.Replace(content, delegate(Match m)
                                                    {
                                                        var sb = new StringBuilder("<w:faker><w:t><a href=\"");
                                                        sb.Append(rels[m.Groups["relid"].Value]);
                                                        if (!string.IsNullOrEmpty(m.Groups["anchor"].Value))
                                                            sb.AppendFormat("#{0}", m.Groups["anchor"].Value);
                                                        sb.Append("\" target='_blank'>");
                                                        AppendInnerText(sb, m.Groups["inner"].Value);
                                                        sb.Append("</a></w:t></w:faker>");
                                                        return sb.ToString();
                                                    });
        }

        private static string ReplaceTables(string content)
        {
            return rxTable.Replace(content, TableReplacement);
        }

        private static string TableReplacement(Match tableMatch)
        {
            var sb = new StringBuilder("<w:p><w:faker><w:t><table>" + Environment.NewLine);

            for (var rowMatch = rxTableRow.Match(tableMatch.Groups["inner"].Value); rowMatch.Success; rowMatch = rowMatch.NextMatch())
            {
                sb.Append("<tr>");

                for (var cellMatch = rxTableCell.Match(rowMatch.Groups["inner"].Value); cellMatch.Success; cellMatch = cellMatch.NextMatch())
                {
                    var gridSpanMatch = rxGridSpan.Match(cellMatch.Groups["inner"].Value);

                    sb.AppendFormat("<td{0}>", gridSpanMatch.Success ? string.Format(" colspan=\"{0}\"", gridSpanMatch.Groups["colspan"].Value) : string.Empty);

                    AppendParagraphs(sb, cellMatch.Groups["inner"].Value);
                    //AppendInnerText(sb, cellMatch.Groups["inner"].Value);

                    sb.Append("</td>" + Environment.NewLine);
                }

                sb.Append("</tr>" + Environment.NewLine);
            }

            sb.Append("</table></w:t></w:faker></w:p>" + Environment.NewLine);

            return sb.ToString();
        }

		private string ReplaceImages(string content)
		{
			return rxDrawingImage.Replace(content,
			                              DrawingImageReplacement);
		}

        private static readonly Regex rxImageBlip = new Regex(@"<a:blip\s+r:(?<embedLink>embed|link)=""(?<relid>[^""]+)""", RegexOptions.Compiled | RegexOptions.Singleline);

        private string DrawingImageReplacement(Match drawingMatch)
        {
            string replacement = "<w:faker><w:t>[Word Drawing Removed]</w:t></w:faker>";

            var blipMatch = rxImageBlip.Match(drawingMatch.Groups["inner"].Value);
            if (blipMatch.Success)
            {
                var relKey = blipMatch.Groups["relid"].Value;
                var relValue = rels[relKey];
                //  we link to images whether they were originally linked or embedded. code in LoadRels copies embedded images to local files.
                replacement = string.Format("<w:faker><w:t><img src=\"{0}\" /></w:t></w:faker>", relValue);
            }

            return replacement;
        }

        private static string ReadAllPart(Package p, string whatPart)
        {
            string s;
            var pp = p.GetPart(new Uri(whatPart, UriKind.Relative));
            using (var pps = pp.GetStream())
            {
                using (var sr = new StreamReader(pps))
                    s = sr.ReadToEnd();

                pps.Close();
            }
            return s;
        }
    }
}
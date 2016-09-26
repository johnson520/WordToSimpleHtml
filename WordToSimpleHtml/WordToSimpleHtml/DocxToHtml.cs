using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;

namespace WordToSimpleHtml
{
    public class DocxToHtml
    {
        public delegate void ErrorLogger(string s);

        private const string AspxPrefix =
            @"<%@ Page Title=""Trickster Cards {0}"" Language=""C#"" MasterPageFile=""~/home/home.master"" CodeBehind=""~/home/InAppPage.cs"" Inherits=""Hearts.home.InAppPage"" AutoEventWireup=""true"" %>
<asp:Content runat=""server"" ContentPlaceHolderID=""mainBody"">
<div class=""main-body-content"">
";

        private const string AspxSuffix = @"</div>
</asp:Content>
";

        private static readonly Regex rxRelationship =
            new Regex(@"<Relationship\s+Id=""(?<key>[^""]+)"".+?Type=""(?<type>[^""]+)"".+?Target=""(?<value>[^""]+)""(?:\s+TargetMode=""(?<mode>[^""]+)"")?.*?/>",
                RegexOptions.Singleline | RegexOptions.Compiled);

        private static readonly Regex rxBody = new Regex(@"<w:body\b[^>]*>(?<inner>.+?)</w:body>", RegexOptions.Singleline | RegexOptions.Compiled);
        private static readonly Regex rxParagraph = new Regex(@"<w:p\b[^>]*>(?<inner>.+?)</w:p>", RegexOptions.Singleline | RegexOptions.Compiled);
        private static readonly Regex rxText = new Regex(@"<w:(?<br>br)/>|<w:(?<tab>tab)/>|<w:t\b[^>]*>(?<inner>.+?)</w:t>", RegexOptions.Singleline | RegexOptions.Compiled);
        private static readonly Regex rxRun = new Regex(@"<(?<tag>w:(?:fake)?r)\b[^>]*>(?<inner>.+?)</\k<tag>>", RegexOptions.Singleline | RegexOptions.Compiled);
        private static readonly Regex rxRunProp = new Regex(@"<w:rPr\b[^>]*>(?<inner>.+?)</w:rPr>", RegexOptions.Singleline | RegexOptions.Compiled);
        private static readonly Regex rxPStyle = new Regex(@"<w:pStyle\s+w:val=""(?<style>[^""]+)""\s*/>", RegexOptions.Singleline | RegexOptions.Compiled);
        private static readonly Regex rxHeadingStyle = new Regex(@"^Heading(?<n>\d+)$", RegexOptions.Singleline | RegexOptions.Compiled);

        private static readonly Regex rxHyperlink = new Regex(@"<w:hyperlink\s+r:id=""(?<relid>[^""]+)""(?:\s+w:anchor=""(?<anchor>[^""]+)"")?[^>]*>(?<inner>.+?)</w:hyperlink>",
            RegexOptions.Singleline | RegexOptions.Compiled);

        private static readonly Regex rxDrawingImage = new Regex("<w:drawing>(?<inner>.+?)</w:drawing>",
            RegexOptions.Singleline | RegexOptions.Compiled);

        private static readonly Regex rxTable = new Regex("<w:tbl>(?<inner>.+?)</w:tbl>", RegexOptions.Singleline | RegexOptions.Compiled);
        private static readonly Regex rxTableStyle = new Regex(@"<w:tblStyle\s+w:val=""(?<styleName>[^""]+)""\s*/>");
        private static readonly Regex rxTableRow = new Regex(@"<w:tr\b[^>]*>(?<inner>.+?)</w:tr>", RegexOptions.Singleline | RegexOptions.Compiled);
        private static readonly Regex rxTableCell = new Regex("<w:tc>(?<inner>.+?)</w:tc>", RegexOptions.Singleline | RegexOptions.Compiled);
        private static readonly Regex rxGridSpan = new Regex(@"<w:gridSpan\s+w:val=""(?<colspan>\d+)""/>");
        private static readonly Regex rxTableCleanup = new Regex(@"<p>\s*(?<tag><table[^>]*>)", RegexOptions.Compiled | RegexOptions.Singleline);
        private static readonly Regex rxTableEndCleanup = new Regex(@"</table>\s*</p>", RegexOptions.Compiled | RegexOptions.Singleline);
        private static readonly Regex rxInnerText = new Regex(">(?<inner>[^>]*)<", RegexOptions.Compiled | RegexOptions.Singleline);
        private static readonly Regex rxWebsomething = new Regex(@"\b[wW]eb(?<something>(?:site|page)s?)\b", RegexOptions.Compiled | RegexOptions.Singleline);
        private static readonly Regex rxWeb = new Regex(@"\bweb\b", RegexOptions.Compiled | RegexOptions.Singleline);
        private static readonly Regex rxInternet = new Regex(@"\binternet\b", RegexOptions.Compiled | RegexOptions.Singleline);
        private static readonly Regex rxQuotePunctuation = new Regex(@"”(?<punc>[,\.])", RegexOptions.Compiled | RegexOptions.Singleline);
        private static readonly Regex rxTestDrive = new Regex(@"\btestdrive\b", RegexOptions.IgnoreCase | RegexOptions.Compiled | RegexOptions.Singleline);
        private static readonly Regex rxLonelyPsInTds = new Regex(@"<td>\s*<p(?<pclass>\s+class=""[^""]+"")\s*>(?<inner>(?:(?!<p>).)*?)</p>\s*</td>", RegexOptions.Compiled | RegexOptions.Singleline);
        private static readonly Regex rxEmptyPs = new Regex(@"<p[^>]*>\s*</p>\s*", RegexOptions.Compiled | RegexOptions.Singleline);
        private static readonly Regex rxTitleP = new Regex(@"^\s*<p\s+class=""Title"">\s*(?<inner>.*?)\s*</p>\s*", RegexOptions.Compiled | RegexOptions.Singleline);
        private static readonly Regex rxInitialH1 = new Regex(@"<h1[^>]*>\s*(?<inner>.*?)\s*</h1>\s*", RegexOptions.Compiled | RegexOptions.Singleline);

        private static readonly Regex rxImgInP = new Regex(@"<p>\s*(?:<[bi]>)*\s*<img\s+src=""(?<src>[^""]+)""\s*/>\s*(?:<br\s?/>)*\s*(?<caption>.*?)\s*(?:</[bi]>\s*)*</p>",
            RegexOptions.Compiled | RegexOptions.Singleline);

        private static readonly Regex rxNoBI = new Regex("</?[bi]>", RegexOptions.Compiled);
        private static readonly Regex rxNoTags = new Regex("<[^>]*>", RegexOptions.Compiled);
        private static readonly Regex rxUnneededBr = new Regex(@"(?<keep><p>)(?:\s*<br\s*/>)+|(?:<br\s*/>\s*)+(?<keep></p>)");
        private static readonly Regex rxListAfterP = new Regex(@"<p>(?<inner>(?:(?!:?</p>).)*):</p>\s*<ul>");
        private static readonly Regex rxImageBlip = new Regex(@"<a:blip\s+r:(?<embedLink>embed|link)=""(?<relid>[^""]+)""", RegexOptions.Compiled | RegexOptions.Singleline);
        private static readonly Regex rxAbsoluteUrl = new Regex("(?:https?:)?//", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex rxFontAwesome = new Regex(@"<i>\s*(?<faclass>fa-\w+)\s*</i>", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        private readonly ErrorLogger logError;
        private readonly Dictionary<string, string> rels = new Dictionary<string, string>();
        private string htmlDir;
        private readonly bool makeAccordion;

        public DocxToHtml(ErrorLogger logger, List<string> options)
        {
            logError = logger;
            makeAccordion = options.Contains("-accordion") || options.Contains("-a");
        }

        public void Convert(string docxFile, string htmlFile, string imageFilePrefix, out string foundTitle)
        {
            foundTitle = string.Empty;

            if (!File.Exists(docxFile))
            {
                logError($"Word file \"{docxFile}\" does not exist.");
                return;
            }

            htmlDir = Path.GetDirectoryName(htmlFile);
            if (htmlDir == null || !Directory.Exists(htmlDir))
            {
                logError($"HTML file directory \"{htmlDir ?? "null"}\" does not exist");
                return;
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

                if (string.Compare(Path.GetExtension(htmlFile), ".aspx", StringComparison.InvariantCulture) == 0)
                {
                    html = $"{string.Format(AspxPrefix, foundTitle)}{html}{AspxSuffix}";
                }

                File.WriteAllText(htmlFile, html, Encoding.UTF8);
            }
            catch (Exception theE)
            {
                File.WriteAllText(htmlFile, $"Exception occurred converting docx to html: {theE.Message} at {theE.StackTrace}");
                throw;
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

                    if (style.Contains("Normal") || style == "BodyText")
                    {
                        style = string.Empty;
                    }
                    else if (style == "ListParagraph" || style == "ListBullet")
                    {
                        if (!inList)
                        {
                            //	to figure out number vs. bullet, we have to redirect into numbering.xml and lookup the <w:numId w:val="1"/>.
                            //	<w:ilvl w:val="0"/> could tell us nestedness.
                            //listTag = rxNumberList.IsMatch(pInner) ? "ol" : "ul";
                            listTag = "ul";
                            sb.AppendLine($"<{listTag}>");
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
                    sb.AppendLine($"</{listTag}>");
                    inList = false;
                }

                //  do some work to avoid outputting empty elements (empty <ul></ul> will still be output)
                var preTagLocation = sb.Length;

                var classAttribute = string.IsNullOrEmpty(style) ? string.Empty : $" class=\"{style}\"";
                sb.Append($"<{tag}{classAttribute}>");

                var preTextLocation = sb.Length;

                AppendInnerText(sb, pInner);

                if (sb.Length == preTextLocation)
                    sb.Remove(preTagLocation, sb.Length - preTagLocation);
                else
                    sb.AppendFormat("</{0}>" + Environment.NewLine, tag);
            }

            if (inList)
            {
                sb.AppendLine($"</{listTag}>");
            }
        }

        private static string CollectParagraphs(string bodyContent)
        {
            var sb = new StringBuilder();
            AppendParagraphs(sb, bodyContent);

            return sb.ToString();
        }

        private static string FinalCleanup(string bodyContent, out string foundTitle)
        {
            bodyContent = rxTableCleanup.Replace(bodyContent, "<div class=\"table-wrapper\">${tag}");
            bodyContent = rxTableEndCleanup.Replace(bodyContent, "</table></div>");

            bodyContent = rxLonelyPsInTds.Replace(bodyContent, "<td${pclass}>${inner}</td>");

            bodyContent = rxImgInP.Replace(bodyContent, m =>
                string.Format(
                    "<p class=\"img-in-p\"><span class=\"img-wrapper\"><img alt=\"{2}\" src=\"/home/help/{1}?ver={3:MMdd}\" /><br />{0}</span></p>",
                    rxNoBI.Replace(m.Groups["caption"].Value, string.Empty), m.Groups["src"].Value, rxNoTags.Replace(m.Groups["caption"].Value, string.Empty), DateTime.UtcNow));

            bodyContent = rxUnneededBr.Replace(bodyContent, "${keep}");
            bodyContent = rxEmptyPs.Replace(bodyContent, string.Empty);

            //  there are two ways we might find a title in the word document: an initial title paragraph or an initial h1
            var titleMatch = rxTitleP.Match(bodyContent);
            if (titleMatch.Success)
            {
                foundTitle = titleMatch.Groups["inner"].Value;
                bodyContent = rxTitleP.Replace(bodyContent, "<h1 class=\"title\">${inner}</h1>" + Environment.NewLine);
            }
            else
            {
                foundTitle = null;
            }

            var h1Match = rxInitialH1.Match(bodyContent);
            if (h1Match.Success)
            {
                foundTitle = h1Match.Groups["inner"].Value;
                bodyContent = rxInitialH1.Replace(bodyContent, "<h1 class=\"title\">${inner}</h1>" + Environment.NewLine);
            }

            bodyContent = rxListAfterP.Replace(bodyContent, "<p class=\"before-help-ul\">${inner}:</p>" + Environment.NewLine + "<ul class=\"help-ul\">");

            bodyContent = rxFontAwesome.Replace(bodyContent, "<i class=\"fa ${faclass}\"></i>");

            return bodyContent;
        }

        private static string ImageFileName(string imageFilePrefix, string relValue)
        {
            var name = imageFilePrefix + Path.GetFileName(relValue);

            //var uniqueSuffix = 0;
            //while (File.Exists(htmlDir + name))
            //    name = string.Format("{0}{1}-{2}{3}", imageFilePrefix, Path.GetFileNameWithoutExtension(relValue), ++uniqueSuffix, Path.GetExtension(relValue));

            return name;
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

        private static string ReplaceTables(string content)
        {
            return rxTable.Replace(content, TableReplacement);
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

        private static string TableReplacement(Match tableMatch)
        {
            var classAttribute = string.Empty;
            var styleMatch = rxTableStyle.Match(tableMatch.Groups["inner"].Value);
            if (styleMatch.Success)
            {
                classAttribute = $" class=\"{styleMatch.Groups["styleName"].Value}\"";
            }

            var sb = new StringBuilder($"<w:p><w:faker><w:t><table{classAttribute}>{Environment.NewLine}");

            for (var rowMatch = rxTableRow.Match(tableMatch.Groups["inner"].Value); rowMatch.Success; rowMatch = rowMatch.NextMatch())
            {
                sb.Append("<tr>");

                for (var cellMatch = rxTableCell.Match(rowMatch.Groups["inner"].Value); cellMatch.Success; cellMatch = cellMatch.NextMatch())
                {
                    var gridSpanMatch = rxGridSpan.Match(cellMatch.Groups["inner"].Value);

                    var spanAttribute = gridSpanMatch.Success ? $" colspan=\"{gridSpanMatch.Groups["colspan"].Value}\"" : string.Empty;
                    sb.Append($"<td{spanAttribute}>");

                    AppendParagraphs(sb, cellMatch.Groups["inner"].Value);
                    //AppendInnerText(sb, cellMatch.Groups["inner"].Value);

                    sb.Append("</td>" + Environment.NewLine);
                }

                sb.Append("</tr>" + Environment.NewLine);
            }

            sb.Append("</table></w:t></w:faker></w:p>" + Environment.NewLine);

            return sb.ToString();
        }

        private void AppendDataTitle(StringBuilder sb, string href, string fallBackTitle)
        {
            fallBackTitle = Regex.Replace(fallBackTitle, @"\b[a-z]", m => m.Value.ToUpper());

            var path = htmlDir + href;
            if (File.Exists(path))
            {
                var h1Match = rxInitialH1.Match(File.ReadAllText(path));
                sb.Append(h1Match.Success ? $" data-title=\"{h1Match.Groups["inner"].Value}\"" : $" data-title=\"{fallBackTitle}\"");
            }
            else
            {
                sb.Append($" data-title=\"{fallBackTitle}\"");
            }
        }

        private string ConvertContentToHtml(string content, out string foundTitle)
        {
            foundTitle = string.Empty;

            var bodyMatch = rxBody.Match(content);
            if (!bodyMatch.Success)
                return string.Empty;

            var bodyContent = bodyMatch.Groups["inner"].Value;

            bodyContent = ReplaceImages(bodyContent);
            bodyContent = ReplaceHyperlinks(bodyContent);
            bodyContent = ReplaceTables(bodyContent);

            bodyContent = CollectParagraphs(bodyContent);
            bodyContent = FinalCleanup(bodyContent, out foundTitle);

            if (makeAccordion)
                bodyContent = MakeAccordionOnH2(bodyContent);

            return bodyContent;
        }

        private static readonly Regex rxH2AndFollowing = new Regex(@"<h2\s*(?<h2attr>[^>]*)>(?<h2inner>.*?)</h2>(?<following>.*?(([\r\n]+.*?)*)*)(?=<h2|$)", RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.Singleline);
        private static readonly Regex rxNonWord = new Regex(@"\W+", RegexOptions.Compiled | RegexOptions.Singleline);
        private const string accordionStyle = @"<style>
    h2[accordion] {
        cursor: pointer;
        padding-left: 20px !important;
        position: relative;
        margin-bottom: 0.5em !important;
    }
    h2[accordion].acc-open {
        margin-bottom: 0 !important;
    }
    h2[accordion]::before {
        font-family: FontAwesome;
        content: ""\f0da"";
        position: absolute;
        left: 0;
        top: 3px;
    }
    h2[accordion].acc-open::before {
        content: ""\f0d7"";
    }
    div[accordion-body] {
        padding-left: 20px;
    }
    div[accordion-body].acc-body-closed {
        display: none;
    }
</style>
";

        private const string accordionScript = @"<script>
    function toggleAccordion(accId) {
        var accH2 = document.querySelector(""h2[accordion='"" + accId + ""']"");
        var accBody = document.querySelector(""div[accordion-body='"" + accId + ""']"");
        if (accH2.classList.contains(""acc-closed"")) {
            accH2.classList.remove(""acc-closed"");
            accH2.classList.add(""acc-open"");
            accBody.classList.remove(""acc-body-closed"");
            accBody.classList.add(""acc-body-open"");
        } else {
            accH2.classList.remove(""acc-open"");
            accH2.classList.add(""acc-closed"");
            accBody.classList.remove(""acc-body-open"");
            accBody.classList.add(""acc-body-closed"");
        }
    }
</script>
";

        private static string MakeAccordionOnH2(string bodyContent)
        {
            bodyContent = rxH2AndFollowing.Replace(bodyContent, m =>
            {
                var id = rxNonWord.Replace(rxNoTags.Replace(m.Groups["h2inner"].Value, string.Empty), string.Empty).ToLowerInvariant();
                if (id.Length > 32)
                    id = id.Substring(0, 32);

                return $"<h2 accordion='{id}' {m.Groups["h2attr"].Value} class='acc-closed' onclick='toggleAccordion(\"{id}\")'>{m.Groups["h2inner"].Value}</h2>\n<div accordion-body='{id}' class='acc-body-closed'>{m.Groups["following"].Value}</div>\n";
            });

            return $"{accordionStyle}{bodyContent}{accordionScript}";
        }

        private string DrawingImageReplacement(Match drawingMatch)
        {
            var replacement = "<w:faker><w:t>[Word Drawing Removed]</w:t></w:faker>";

            var blipMatch = rxImageBlip.Match(drawingMatch.Groups["inner"].Value);
            if (blipMatch.Success)
            {
                var relKey = blipMatch.Groups["relid"].Value;
                var relValue = rels[relKey];
                //  we link to images whether they were originally linked or embedded. code in LoadRels copies embedded images to local files.
                replacement = $"<w:faker><w:t><img src=\"{relValue}\" /></w:t></w:faker>";
            }

            return replacement;
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

        private string ReplaceHyperlinks(string content)
        {
            return rxHyperlink.Replace(content, delegate(Match m)
            {
                var href = rels[m.Groups["relid"].Value];
                var isMailTo = href.Contains("@");
                var hash = m.Groups["anchor"].Value;

                var sb = new StringBuilder("<w:faker><w:t>");
                if (isMailTo)
                {
                    if (href.StartsWith("mailto:"))
                        href = href.Substring("mailto:".Length);
                    sb.Append($"<a href=\"mailto:{href}\"");
                }
                else
                    sb.AppendFormat("<a href=\"{0}{1}\"", href, !string.IsNullOrEmpty(hash) ? $"#{hash}" : string.Empty);

                if (rxAbsoluteUrl.IsMatch(href))
                    sb.Append(" target=\"_blank\"");
                else if (!isMailTo)
                    AppendDataTitle(sb, href, rxText.Match(m.Groups["inner"].Value).Groups["inner"].Value);

                sb.Append(">");
                AppendInnerText(sb, m.Groups["inner"].Value);
                sb.Append("</a></w:t></w:faker>");
                return sb.ToString();
            });
        }

        private string ReplaceImages(string content)
        {
            return rxDrawingImage.Replace(content, DrawingImageReplacement);
        }
    }
}
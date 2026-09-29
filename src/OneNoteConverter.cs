using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Xml;

// Converts OneNote 2013 page XML into Markdown.
public sealed class OneNoteConverter
{
    public const string NS = "http://schemas.microsoft.com/office/onenote/2013/onenote";

    readonly string _assetsDir;          // where images are written (may be null => skip)
    readonly Func<string, int, string> _assetNamer;
    int _imageCounter;

    // Maps quickStyleIndex -> style name (e.g. "h2", "p", "cite")
    readonly Dictionary<int, string> _styleNames = new Dictionary<int, string>();

    public OneNoteConverter(string assetsDir, Func<string, int, string> assetNamer)
    {
        _assetsDir = assetsDir;
        _assetNamer = assetNamer;
    }

    public string ConvertPage(string pageXml, string pageTitle)
    {
        XmlDocument doc = new XmlDocument();
        doc.PreserveWhitespace = false;
        doc.LoadXml(pageXml);

        XmlNamespaceManager ns = new XmlNamespaceManager(doc.NameTable);
        ns.AddNamespace("one", NS);

        LoadStyleNames(doc, ns);

        StringBuilder md = new StringBuilder();
        md.Append("# ").Append(EscapeTitle(pageTitle)).Append("\n\n");

        // Page title as authored inside the page (may differ from hierarchy name).
        XmlNode titleNode = doc.SelectSingleNode("/one:Page/one:Title", ns);
        if (titleNode != null)
        {
            string inner = PlainText(titleNode);
            if (!string.IsNullOrEmpty(inner) && inner.Trim() != pageTitle.Trim())
                md.Append("**").Append(inner.Trim()).Append("**\n\n");
        }

        foreach (XmlNode outline in doc.SelectNodes("/one:Page/one:Outline", ns))
            WriteOutline(outline, md, ns, 0);

        return md.ToString();
    }

    void LoadStyleNames(XmlDocument doc, XmlNamespaceManager ns)
    {
        foreach (XmlNode q in doc.SelectNodes("/one:Page/one:QuickStyleDef", ns))
        {
            string idx = TextRunExtractor.Attr(q, "index");
            string nm = TextRunExtractor.Attr(q, "name");
            int i;
            if (idx != null && nm != null && int.TryParse(idx, out i))
                _styleNames[i] = nm;
        }
    }

    // <one:Outline> == one bullet level. Depending on the schema revision the
    // actual <one:OE> paragraphs hang directly off the Outline or off an
    // intermediate <one:OEChildren>, so both shapes are handled here.
    void WriteOutline(XmlNode outline, StringBuilder md, XmlNamespaceManager ns, int depth)
    {
        foreach (XmlNode child in outline.ChildNodes)
        {
            if (child.NodeType != XmlNodeType.Element) continue;

            switch (child.LocalName)
            {
                case "OE":
                    WriteOE(child, md, ns, depth);
                    break;

                case "OEChildren":
                    foreach (XmlNode oe in child.SelectNodes("one:OE", ns))
                        WriteOE(oe, md, ns, depth);
                    foreach (XmlNode nested in child.SelectNodes("one:Outline", ns))
                        WriteOutline(nested, md, ns, depth);
                    break;

                case "Outline":
                    WriteOutline(child, md, ns, depth);
                    break;
            }
        }
    }

    void WriteOE(XmlNode oe, StringBuilder md, XmlNamespaceManager ns, int depth)
    {
        string quickStyle = TextRunExtractor.Attr(oe, "quickStyleIndex");
        string styleName = null;
        int qi;
        if (quickStyle != null && int.TryParse(quickStyle, out qi) && _styleNames.ContainsKey(qi))
            styleName = _styleNames[qi];

        int level = HeadingLevel(styleName);
        bool isBullet = oe.SelectSingleNode("one:List", ns) != null;

        // Indent for nested content
        string pad = new string(' ', depth * 2);

        if (level > 0)
        {
            string text = InlineContent(oe, ns);
            if (text.Trim().Length > 0)
            {
                md.Append(pad);
                for (int h = 0; h < level; h++) md.Append("#");
                md.Append(" ").Append(text.Trim()).Append("\n\n");
            }
        }
        else if (isBullet)
        {
            string text = InlineContent(oe, ns);
            if (text.Trim().Length > 0)
                md.Append(pad).Append("- ").Append(text.Trim()).Append("\n");
        }
        else
        {
            string text = InlineContent(oe, ns);
            if (text.Trim().Length > 0)
                md.Append(text.TrimEnd()).Append("\n\n");
        }

        // Tables rendered by OneNote arrive as an HTML blob.
        XmlNode table = oe.SelectSingleNode("one:OEChildren/one:T[@tableHTML] | one:T[@tableHTML]", ns);
        if (table != null)
        {
            string html = TextRunExtractor.Attr(table, "tableHTML");
            if (!string.IsNullOrEmpty(html))
            {
                md.Append(ConvertTableHtml(html)).Append("\n\n");
            }
        }

        // Recurse into children (sub-bullets, and anything else)
        XmlNode kids = oe.SelectSingleNode("one:OEChildren", ns);
        if (kids != null)
        {
            foreach (XmlNode childOE in kids.SelectNodes("one:OE", ns))
                WriteOE(childOE, md, ns, depth + 1);
            foreach (XmlNode childOutline in kids.SelectNodes("one:Outline", ns))
                WriteOutline(childOutline, md, ns, depth + 1);
        }
    }

    static int HeadingLevel(string styleName)
    {
        if (string.IsNullOrEmpty(styleName)) return 0;
        if (styleName == "PageTitle") return 0;         // title handled separately
        if (styleName.Length == 2 && styleName[0] == 'h' && styleName[1] >= '1' && styleName[1] <= '6')
            return styleName[1] - '0';
        return 0;
    }

    // Renders the inline content of an <one:OE> as a single line of Markdown.
    string InlineContent(XmlNode oe, XmlNamespaceManager ns)
    {
        StringBuilder sb = new StringBuilder();

        foreach (XmlNode node in oe.ChildNodes)
        {
            if (node.NodeType != XmlNodeType.Element) continue;

            switch (node.LocalName)
            {
                case "T":
                    {
                        string raw = InnerTextOf(node);
                        if (LooksLikeHtml(raw))
                            sb.Append(ConvertInlineHtml(raw));
                        else
                        {
                            var tr = new TextRunExtractor();
                            tr.Extract(node);
                            sb.Append(tr.Out.ToString());
                        }
                        break;
                    }
                case "Image":
                    sb.Append(RenderImage(node, ns));
                    break;
                case "OEChildren":
                    break; // handled by caller
            }
        }

        return sb.ToString();
    }

    string InnerTextOf(XmlNode t)
    {
        if (t == null) return "";
        StringBuilder sb = new StringBuilder();
        foreach (XmlNode c in t.ChildNodes)
        {
            if (c.NodeType == XmlNodeType.CDATA || c.NodeType == XmlNodeType.Text)
                sb.Append(c.Value ?? "");
            else if (c.LocalName == "r")
            {
                foreach (XmlNode rc in c.ChildNodes)
                    if (rc.NodeType == XmlNodeType.CDATA || rc.NodeType == XmlNodeType.Text)
                        sb.Append(rc.Value ?? "");
            }
        }
        // CDATA holds text verbatim, so HTML entities inside it are still encoded.
        return DecodeEntities(sb.ToString());
    }

    /// <summary>
    /// Resolves the HTML/XML entities OneNote leaves inside CDATA text
    /// (&amp;nbsp;, &amp;#xA; and friends). "&amp;amp;" is resolved last so that
    /// a literal "&amp;amp;lt;" cannot collapse all the way down to "&lt;".
    /// </summary>
    public static string DecodeEntities(string s)
    {
        if (string.IsNullOrEmpty(s) || s.IndexOf('&') < 0) return s;

        StringBuilder sb = new StringBuilder(s.Length);
        for (int i = 0; i < s.Length; i++)
        {
            if (s[i] != '&') { sb.Append(s[i]); continue; }

            int semi = s.IndexOf(';', i + 1);
            if (semi < 0 || semi - i > 10) { sb.Append('&'); continue; }

            string ent = s.Substring(i + 1, semi - i - 1);
            string rep = null;

            if (ent.Length > 1 && ent[0] == '#')
            {
                int cp;
                bool ok = ent[1] == 'x' || ent[1] == 'X'
                    ? int.TryParse(ent.Substring(2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out cp)
                    : int.TryParse(ent.Substring(1), out cp);
                if (ok && cp > 0 && cp <= 0x10FFFF)
                {
                    // Soft line breaks inside a paragraph become spaces.
                    if (cp == 0x0A || cp == 0x0D) sb.Append(' ');
                    else sb.Append(char.ConvertFromUtf32(cp));
                    i = semi;
                    continue;
                }
            }
            else
            {
                switch (ent)
                {
                    case "amp": rep = "&"; break;
                    case "lt": rep = "<"; break;
                    case "gt": rep = ">"; break;
                    case "quot": rep = "\""; break;
                    case "apos": rep = "'"; break;
                    case "nbsp": rep = " "; break;
                    case "hellip": rep = "\u2026"; break;
                    case "ndash": rep = "\u2013"; break;
                    case "mdash": rep = "\u2014"; break;
                    case "bull": rep = "\u2022"; break;
                    case "auml": rep = "\u00e4"; break;
                    case "ouml": rep = "\u00f6"; break;
                    case "uuml": rep = "\u00fc"; break;
                    case "Auml": rep = "\u00c4"; break;
                    case "Ouml": rep = "\u00d6"; break;
                    case "Uuml": rep = "\u00dc"; break;
                    case "szlig": rep = "\u00df"; break;
                    case "euro": rep = "\u20ac"; break;
                    case "copy": rep = "\u00a9"; break;
                    case "reg": rep = "\u00ae"; break;
                    case "deg": rep = "\u00b0"; break;
                    case "plusmn": rep = "\u00b1"; break;
                    case "times": rep = "\u00d7"; break;
                    case "middot": rep = "\u00b7"; break;
                    case "laquo": rep = "\u00ab"; break;
                    case "raquo": rep = "\u00bb"; break;
                    case "ldquo": rep = "\u201c"; break;
                    case "rdquo": rep = "\u201d"; break;
                    case "lsquo": rep = "\u2018"; break;
                    case "rsquo": rep = "\u2019"; break;
                    case "eacute": rep = "\u00e9"; break;
                    case "egrave": rep = "\u00e8"; break;
                    case "agrave": rep = "\u00e0"; break;
                }
            }

            if (rep != null) { sb.Append(rep); i = semi; }
            else sb.Append('&');
        }
        return sb.ToString();
    }

    // OneNote frequently stores a whole <one:T> as a literal HTML fragment that
    // mixes plain text and markup, e.g. "Zielgruppe: <span ...>Auszubildende</span>."
    // Such fragments are common enough that sniffing for any inline tag (not just
    // one at the very start) is the reliable test.
    static readonly string[] HtmlTags =
    {
        "<span", "<b>", "<b ", "<i>", "<i ", "<em", "<font", "<a ", "<a>",
        "<br", "<code", "<sub", "<sup", "<u>", "<strong", "<s>", "<strike", "<img"
    };

    static bool LooksLikeHtml(string s)
    {
        if (string.IsNullOrEmpty(s)) return false;
        for (int i = 0; i < s.Length; i++)
        {
            if (s[i] != '<') continue;
            for (int t = 0; t < HtmlTags.Length; t++)
            {
                string tag = HtmlTags[t];
                if (i + tag.Length <= s.Length &&
                    string.Compare(s, i, tag, 0, tag.Length, StringComparison.OrdinalIgnoreCase) == 0)
                    return true;
            }
        }
        return false;
    }

    string RenderImage(XmlNode img, XmlNamespaceManager ns)
    {
        string alt = TextRunExtractor.Attr(img, "alt");
        string format = TextRunExtractor.Attr(img, "format");
        if (string.IsNullOrEmpty(format)) format = "png";

        XmlNode data = img.SelectSingleNode("one:Data", ns);
        if (data == null) return "";

        string b64 = data.InnerText;
        if (string.IsNullOrEmpty(b64)) return "";

        if (_assetsDir == null) return "";   // images suppressed

        try
        {
            byte[] bytes = Convert.FromBase64String(b64);
            _imageCounter++;
            string fname = _assetNamer(format, _imageCounter);
            Directory.CreateDirectory(_assetsDir);
            File.WriteAllBytes(Path.Combine(_assetsDir, fname), bytes);
            return "![" + EscapeAlt(ShortAlt(alt)) + "](assets/" + fname + ")";
        }
        catch (Exception)
        {
            return "";
        }
    }

    /// <summary>
    /// OneNote stores the full OCR transcript in the image's alt attribute,
    /// which would otherwise become an unreadable multi-kilobyte alt text.
    /// </summary>
    static string ShortAlt(string s)
    {
        if (string.IsNullOrWhiteSpace(s)) return "image";
        string t = Collapse(s);
        const int max = 120;
        if (t.Length <= max) return t;
        int cut = t.LastIndexOf(' ', max);
        if (cut < 40) cut = max;
        return t.Substring(0, cut).TrimEnd() + "\u2026";
    }

    static string Collapse(string s)
    {
        return s.Replace("\r", " ").Replace("\n", " ").Replace("\u00A0", " ").Trim();
    }

    static string EscapeAlt(string s)
    {
        return s.Replace("[", "\\[").Replace("]", "\\]");
    }

    static string EscapeTitle(string s)
    {
        if (string.IsNullOrEmpty(s)) return "Untitled";
        return s.Replace("\r", " ").Replace("\n", " ").Trim();
    }

    static string PlainText(XmlNode n)
    {
        if (n == null) return "";
        StringBuilder sb = new StringBuilder();
        foreach (XmlNode c in n.ChildNodes)
        {
            if (c.NodeType == XmlNodeType.CDATA || c.NodeType == XmlNodeType.Text) sb.Append(c.Value ?? "");
            else if (c.LocalName == "r" || c.LocalName == "T")
                sb.Append(PlainText(c));
        }
        return sb.ToString();
    }

    // OneNote stores some formatted runs as literal HTML inside <one:T>.
    // Those fragments are frequently *not* well-formed XML (unclosed spans,
    // stray entities), so this is a tolerant scanner rather than a DOM parse:
    // it never throws and degrades to plain text instead of losing content.
    string ConvertInlineHtml(string html)
    {
        if (string.IsNullOrEmpty(html)) return "";

        StringBuilder sb = new StringBuilder(html.Length);

        // Marks that were opened and still need to be closed at the end.
        List<string> open = new List<string>();
        List<string> links = new List<string>();   // hrefs, aligned with <a>
        string openSpan = null;                    // currently active <span> mark

        int i = 0;
        while (i < html.Length)
        {
            char c = html[i];
            if (c != '<') { sb.Append(c); i++; continue; }

            int gt = html.IndexOf('>', i + 1);
            if (gt < 0) { sb.Append(html.Substring(i)); break; }

            string tag = html.Substring(i + 1, gt - i - 1);
            i = gt + 1;

            if (tag.Length == 0) continue;
            if (tag[0] == '!') continue;                       // comment / doctype

            bool closing = tag[0] == '/';
            if (closing) tag = tag.Substring(1);

            int sp = tag.IndexOfAny(new char[] { ' ', '\t', '\r', '\n', '/' });
            string name = (sp < 0 ? tag : tag.Substring(0, sp)).ToLowerInvariant();
            string attrs = sp < 0 ? "" : tag.Substring(sp);

            switch (name)
            {
                case "br":
                    sb.Append("  \n");
                    break;

                case "b":
                case "strong":
                    ApplyMark(sb, open, "**", closing);
                    break;

                case "i":
                case "em":
                    ApplyMark(sb, open, "*", closing);
                    break;

                case "u":
                case "ins":
                    ApplyMark(sb, open, "<u>", closing);
                    break;

                case "s":
                case "strike":
                case "del":
                    ApplyMark(sb, open, "~~", closing);
                    break;

                case "code":
                    ApplyMark(sb, open, "`", closing);
                    break;

                case "span":
                    {
                        // OneNote emits one <span> per formatting run and often
                        // omits </span>, so treat it as a toggle: the previous run
                        // is always closed before the next one begins.
                        if (openSpan != null) { sb.Append(openSpan); openSpan = null; }
                        if (!closing)
                        {
                            string st = (AttrValue(attrs, "style") ?? "").ToLowerInvariant();
                            if (st.Contains("font-weight:bold") || st.Contains("font-weight:700")) { openSpan = "**"; sb.Append("**"); }
                            else if (st.Contains("font-style:italic")) { openSpan = "*"; sb.Append("*"); }
                            else if (st.Contains("text-decoration:underline")) { openSpan = "<u>"; sb.Append("<u>"); }
                        }
                        break;
                    }

                case "a":
                    {
                        if (closing)
                        {
                            if (links.Count > 0)
                            {
                                string href = links[links.Count - 1];
                                links.RemoveAt(links.Count - 1);
                                if (!string.IsNullOrEmpty(href)) sb.Append(" (").Append(href).Append(")");
                            }
                        }
                        else
                        {
                            links.Add(AttrValue(attrs, "href"));
                        }
                        break;
                    }

                case "img":
                    {
                        string src = AttrValue(attrs, "src");
                        string alt = AttrValue(attrs, "alt");
                        if (!string.IsNullOrEmpty(src))
                            sb.Append("![").Append(alt ?? "image").Append("](").Append(src).Append(")");
                        break;
                    }

                default:
                    break;   // unknown tag: drop the tag, keep the text
            }
        }

        if (openSpan != null) { sb.Append(openSpan); openSpan = null; }
        for (int k = open.Count - 1; k >= 0; k--) sb.Append(open[k]);
        return sb.ToString();
    }

    static void ApplyMark(StringBuilder sb, List<string> open, string mark, bool closing)
    {
        if (closing)
        {
            for (int k = open.Count - 1; k >= 0; k--)
            {
                if (open[k] == mark)
                {
                    sb.Append(mark);
                    open.RemoveRange(k, open.Count - k);
                    return;
                }
            }
        }
        else
        {
            open.Add(mark);
            sb.Append(mark);
        }
    }

    /// <summary>Pulls a value out of a raw HTML attribute list.</summary>
    static string AttrValue(string attrs, string name)
    {
        if (string.IsNullOrEmpty(attrs)) return null;
        int at = 0;
        while ((at = attrs.IndexOf(name, at, StringComparison.OrdinalIgnoreCase)) >= 0)
        {
            // must be preceded by whitespace and followed by '='
            if (at == 0 || char.IsWhiteSpace(attrs[at - 1]))
            {
                int eq = at + name.Length;
                while (eq < attrs.Length && char.IsWhiteSpace(attrs[eq])) eq++;
                if (eq < attrs.Length && attrs[eq] == '=')
                {
                    eq++;
                    while (eq < attrs.Length && char.IsWhiteSpace(attrs[eq])) eq++;
                    if (eq < attrs.Length && (attrs[eq] == '"' || attrs[eq] == '\''))
                    {
                        char q = attrs[eq];
                        int end = attrs.IndexOf(q, eq + 1);
                        if (end > eq) return DecodeEntities(attrs.Substring(eq + 1, end - eq - 1));
                    }
                    else
                    {
                        int end = eq;
                        while (end < attrs.Length && !char.IsWhiteSpace(attrs[end]) && attrs[end] != '>') end++;
                        return DecodeEntities(attrs.Substring(eq, end - eq));
                    }
                }
            }
            at += name.Length;
        }
        return null;
    }

    static string StripTags(string s)
    {
        int i = 0;
        StringBuilder sb = new StringBuilder();
        bool inTag = false;
        for (; i < s.Length; i++)
        {
            if (s[i] == '<') inTag = true;
            else if (s[i] == '>') inTag = false;
            else if (!inTag) sb.Append(s[i]);
        }
        return sb.ToString();
    }

    // Converts a OneNote tableHTML blob into a GitHub-flavoured Markdown table.
    string ConvertTableHtml(string html)
    {
        try
        {
            XmlDocument d = new XmlDocument();
            d.LoadXml("<root>" + html + "</root>");

            List<List<string>> rows = new List<List<string>>();
            XmlNodeList trs = d.GetElementsByTagName("tr");
            foreach (XmlNode tr in trs)
            {
                List<string> cells = new List<string>();
                foreach (XmlNode cell in tr.ChildNodes)
                {
                    if (cell.NodeType != XmlNodeType.Element) continue;
                    if (cell.LocalName != "td" && cell.LocalName != "th") continue;
                    // InnerXml keeps inline markup, which the scanner re-reads so
                    // that bold/italic inside cells survives.
                    cells.Add(CleanCell(ConvertInlineHtml(cell.InnerXml)));
                }
                if (cells.Count > 0) rows.Add(cells);
            }

            if (rows.Count == 0) return "";

            int cols = 0;
            foreach (List<string> r in rows) cols = Math.Max(cols, r.Count);
            if (cols == 0) return "";

            StringBuilder sb = new StringBuilder();
            List<string> header = rows[0];
            sb.Append("| ");
            for (int i = 0; i < cols; i++) sb.Append(EscapeCell(i < header.Count ? header[i] : "")).Append(" | ");
            sb.Append("\n| ");
            for (int i = 0; i < cols; i++) sb.Append(" --- | ");
            sb.Append("\n");

            for (int r = 1; r < rows.Count; r++)
            {
                List<string> row = rows[r];
                sb.Append("| ");
                for (int i = 0; i < cols; i++) sb.Append(EscapeCell(i < row.Count ? row[i] : "")).Append(" | ");
                sb.Append("\n");
            }
            return sb.ToString().TrimEnd('\n');
        }
        catch (Exception)
        {
            return "";
        }
    }

    static string CleanCell(string s)
    {
        return s.Replace("\r", " ").Replace("\n", " ").Replace("\u00A0", " ").Trim();
    }

    static string EscapeCell(string s)
    {
        return (s ?? "").Replace("|", "\\|");
    }
}

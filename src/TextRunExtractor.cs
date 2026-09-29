using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

// Extracts inline formatting runs from a OneNote <one:T> node.
public sealed class TextRunExtractor
{
    public StringBuilder Out = new StringBuilder();

    public void Extract(System.Xml.XmlNode tNode)
    {
        // <one:T> holds CDATA (plain text) or child <one:r> runs (formatted).
        foreach (System.Xml.XmlNode child in tNode.ChildNodes)
        {
            if (child.NodeType == System.Xml.XmlNodeType.CDATA ||
                child.NodeType == System.Xml.XmlNodeType.Text)
            {
                Out.Append(OneNoteConverter.DecodeEntities(child.Value ?? ""));
            }
            else if (child.LocalName == "r")
            {
                AppendRun(child);
            }
        }
    }

    void AppendRun(System.Xml.XmlNode run)
    {
        string text = "";
        bool hadBreak = false;
        foreach (System.Xml.XmlNode c in run.ChildNodes)
        {
            if (c.NodeType == System.Xml.XmlNodeType.CDATA || c.NodeType == System.Xml.XmlNodeType.Text)
                text += c.Value ?? "";
            else if (c.LocalName == "br")
                hadBreak = true;
        }

        string prefix = "", suffix = "";
        string style = Attr(run, "style");
        if (style != null)
        {
            if (HasFlag(style, "font-weight:bold") || HasFlag(style, "font-weight:700")) { prefix += "**"; suffix = "**" + suffix; }
            if (HasFlag(style, "font-style:italic")) { prefix = "*" + prefix; suffix = suffix + "*"; }
            if (HasFlag(style, "text-decoration:underline") || HasFlag(style, "text-decoration:line-through")) { prefix = "<u>" + prefix; suffix = suffix + "</u>"; }
        }

        if (text.Length == 0) return;
        text = OneNoteConverter.DecodeEntities(text);
        Out.Append(prefix);
        Out.Append(text);
        Out.Append(suffix);
        if (hadBreak) Out.Append("  \n");
    }

    public static bool HasFlag(string style, string token)
    {
        if (style == null) return false;
        return style.IndexOf(token, StringComparison.OrdinalIgnoreCase) >= 0;
    }

    public static string Attr(System.Xml.XmlNode n, string name)
    {
        if (n == null || n.Attributes == null) return null;
        System.Xml.XmlAttribute a = n.Attributes[name];
        return a == null ? null : a.Value;
    }
}

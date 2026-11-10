// MD Doc Link - shared names and helpers
// Author: MD (Mayank Dholariya)
using System;
using System.Text;

namespace DocLink.Core
{
    public static class Names
    {
        public const string Product = "MD Doc Link";
        public const string Author = "MD (Mayank Dholariya)";
        public const string Version = "1.0.0";
        public const string SheetName = "DocLink";
        /// <summary>Every link this plugin writes starts with this, so it can find and remove its own links.</summary>
        public const string LinkPrefix = "MD | ";
        public const string SetFolder = "MD Doc Link";
        public const string SetLinked = "MD Doc Link - Linked";
        public const string SetDocMissing = "MD Doc Link - Doc missing";
        public const string ItemNameProperty = "Item > Name (display name)";

        public static bool IsMdLinkName(string linkName)
        {
            return linkName != null && linkName.StartsWith(LinkPrefix, StringComparison.Ordinal);
        }

        public static string MdLinkName(string label)
        {
            string l = Clean(label);
            if (l.Length == 0) l = "Document";
            return IsMdLinkName(l) ? l : LinkPrefix + l;
        }

        /// <summary>Single-line, control-char free, max 200 chars.</summary>
        public static string Clean(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return "";
            var sb = new StringBuilder(s.Length);
            bool space = false;
            foreach (char c in s.Trim())
            {
                if (c < 32 || char.IsWhiteSpace(c)) { space = true; continue; }
                if (space && sb.Length > 0) sb.Append(' ');
                space = false;
                sb.Append(c);
            }
            string r = sb.ToString();
            return r.Length > 200 ? r.Substring(0, 200) : r;
        }

        /// <summary>Safe Navisworks Selection Set display name.</summary>
        public static string SanitizeSetName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return "(blank)";
            var sb = new StringBuilder(name.Length);
            foreach (char c in name.Trim())
            {
                if (c < 32) continue;
                if (c == '/' || c == '\\' || c == ':' || c == '*' || c == '?' || c == '"' || c == '<' || c == '>' || c == '|')
                    sb.Append('-');
                else
                    sb.Append(c);
            }
            string s = sb.ToString().Trim();
            if (s.Length == 0) s = "(blank)";
            if (s.Length > 64) s = s.Substring(0, 64).Trim();
            return s;
        }
    }
}

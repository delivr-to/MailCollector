using System;
using System.Collections.Generic;
using System.Net;
using System.Text.RegularExpressions;

namespace MailCollector
{
    // Status values accepted by the delivr.to results upload
    public static class EmailStatus
    {
        public const string Delivered = "Delivered";
        public const string Rewritten = "Rewritten";
        public const string Held = "Held";
    }

    public sealed class LinkResult
    {
        public LinkResult(string status, string link)
        {
            Status = status;
            Link = link;
        }

        // Delivered, Rewritten or Held
        public string Status { get; }

        // The URL to report: the original link when delivered intact, the rewritten
        // href when rewritten, or empty when no link was found
        public string Link { get; }
    }

    // Classifies link payload emails from their HTML so results collected offline
    // match those from the Microsoft 365 and Google integrations.
    public static class LinkClassifier
    {
        // Characters that can only start a new component of a URL. An expected URL
        // followed by one of these still points at the same resource (e.g. the
        // ?xsdata= that Exchange appends to some links).
        const string LinkUrlBoundaryChars = "?&#/";

        static readonly Regex WbrRegex = new Regex(@"<wbr\s*/?>", RegexOptions.IgnoreCase);
        static readonly Regex HrefRegex = new Regex(@"href\s*=\s*[""']([^""']+)[""']", RegexOptions.IgnoreCase);
        static readonly Regex AnchorRegex = new Regex(@"<a\b([^>]*)>(.*?)</a\s*>", RegexOptions.IgnoreCase | RegexOptions.Singleline);
        static readonly Regex TagRegex = new Regex(@"<[^>]+>");
        static readonly Regex DelivrtoLinkRegex = new Regex(
            @"https://(?:dev\.)?delivrto\.me/links/[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}",
            RegexOptions.IgnoreCase);

        // Strips markup that splits or escapes a URL without changing where it points:
        // the <wbr> break opportunities in the visible link text and HTML-escaped '&'.
        public static string NormaliseBody(string html)
        {
            string body = WbrRegex.Replace(html ?? "", "");
            return body.Replace("&amp;", "&");
        }

        // Whether an href opens the resource at url, ignoring appended parameters
        public static bool HrefPointsAt(string href, string url)
        {
            href = (href ?? "").Trim();
            if (string.IsNullOrEmpty(url))
                return false;
            if (href == url)
                return true;
            return href.StartsWith(url, StringComparison.Ordinal)
                && LinkUrlBoundaryChars.IndexOf(href[url.Length]) >= 0;
        }

        public static LinkResult ClassifyLink(string html)
        {
            string body = NormaliseBody(html);
            string expectedUrl = null;
            string anchorHref = null;

            // The template's link anchor shows the URL as its text; the first such
            // anchor after the "requested link" sentence is the payload link.
            int start = body.IndexOf("requested link", StringComparison.OrdinalIgnoreCase);
            foreach (Anchor anchor in GetAnchors(body, Math.Max(start, 0)))
            {
                if (IsAbsoluteHttpUrl(anchor.Text))
                {
                    expectedUrl = anchor.Text;
                    anchorHref = anchor.Href;
                    break;
                }
            }

            // Fall back to a delivr.to-hosted link anywhere in the body
            if (expectedUrl == null)
            {
                Match match = DelivrtoLinkRegex.Match(body);
                if (match.Success)
                    expectedUrl = match.Value;
            }

            if (expectedUrl == null)
                return new LinkResult(EmailStatus.Held, "");

            foreach (Match href in HrefRegex.Matches(body))
            {
                if (HrefPointsAt(href.Groups[1].Value, expectedUrl))
                    return new LinkResult(EmailStatus.Delivered, expectedUrl);
            }

            if (body.Contains(expectedUrl))
                return new LinkResult(EmailStatus.Rewritten, anchorHref ?? "");

            return new LinkResult(EmailStatus.Held, "");
        }

        static bool IsAbsoluteHttpUrl(string text)
        {
            return Uri.TryCreate(text, UriKind.Absolute, out Uri uri)
                && (uri.Scheme == Uri.UriSchemeHttp || uri.Scheme == Uri.UriSchemeHttps)
                && text.IndexOfAny(new[] { ' ', '\t', '\r', '\n' }) < 0;
        }

        static IEnumerable<Anchor> GetAnchors(string body, int start)
        {
            for (Match m = AnchorRegex.Match(body, start); m.Success; m = m.NextMatch())
            {
                string attributes = m.Groups[1].Value;
                Match href = HrefRegex.Match(attributes);
                string text = WebUtility.HtmlDecode(TagRegex.Replace(m.Groups[2].Value, "")).Trim();

                yield return new Anchor(href.Success ? href.Groups[1].Value.Trim() : "", text);
            }
        }

        sealed class Anchor
        {
            public Anchor(string href, string text)
            {
                Href = href;
                Text = text;
            }

            public string Href { get; }
            public string Text { get; }
        }
    }
}

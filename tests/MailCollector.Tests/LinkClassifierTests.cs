using System.Text.RegularExpressions;
using MailCollector;
using Xunit;

namespace MailCollector.Tests
{
    public class LinkClassifierTests
    {
        const string DelivrtoLink = "https://delivrto.me/links/d00dafbf-4b13-4b66-963f-37495b4067ad";
        const string ExternalLink = "https://files.example.com/download/test.js?key=qrstuvwxyz&dl=1";

        // Inserts <wbr> break opportunities into the visible link text, as delivr.to emails do
        static string WrapForDisplay(string url) =>
            Regex.Replace(url, @"([/?#&=_.\-%+])", "$1<wbr>");

        // The link template body, as delivered by Exchange (& in the visible text arrives HTML-escaped)
        static string LinkEmail(string href, string url, string extraAttributes = "") =>
            "<p>Hi,</p><p>Please find your requested link for file: <b>test.doc</b>.</p>" +
            $"<p>Link: <a href=\"{href}\"{extraAttributes} style=\"word-break: break-all; overflow-wrap: anywhere;\">" +
            WrapForDisplay(url).Replace("&", "&amp;") +
            "</a>.</p><p>Email details:<br>- Email ID: <b>c2ab5e1c-b786-4846-99e8-bf7c9163c77a</b></p>";

        static string SafeLinksHref(string url) =>
            "https://gbr01.safelinks.protection.outlook.com/?url=" + System.Uri.EscapeDataString(url) +
            "&amp;data=05%7C02%7Cuser%40example.com%7C0%7C0&amp;sdata=abc%3D&amp;reserved=0";

        [Fact]
        public void UntouchedDelivrtoLinkIsDelivered()
        {
            LinkResult result = LinkClassifier.ClassifyLink(LinkEmail(DelivrtoLink, DelivrtoLink));

            Assert.Equal(EmailStatus.Delivered, result.Status);
            Assert.Equal(DelivrtoLink, result.Link);
        }

        [Fact]
        public void SafeLinksRewriteIsRewrittenAndReportsRewrittenHref()
        {
            string html = LinkEmail(SafeLinksHref(DelivrtoLink), DelivrtoLink, $" originalsrc=\"{DelivrtoLink}\"");

            LinkResult result = LinkClassifier.ClassifyLink(html);

            Assert.Equal(EmailStatus.Rewritten, result.Status);
            Assert.StartsWith("https://gbr01.safelinks.protection.outlook.com/", result.Link);
            Assert.NotEqual(DelivrtoLink, result.Link);
        }

        [Fact]
        public void UntouchedExternalLinkWithEscapedAmpersandIsDelivered()
        {
            string html = LinkEmail(ExternalLink.Replace("&", "&amp;"), ExternalLink);

            LinkResult result = LinkClassifier.ClassifyLink(html);

            Assert.Equal(EmailStatus.Delivered, result.Status);
            Assert.Equal(ExternalLink, result.Link);
        }

        [Fact]
        public void SafeLinksRewriteOfExternalLinkIsRewritten()
        {
            string html = LinkEmail(SafeLinksHref(ExternalLink), ExternalLink, $" originalsrc=\"{ExternalLink.Replace("&", "&amp;")}\"");

            Assert.Equal(EmailStatus.Rewritten, LinkClassifier.ClassifyLink(html).Status);
        }

        [Fact]
        public void ExchangeDecoratedLinkIsDelivered()
        {
            const string shareUrl = "https://files.example.com/s/EaBcDeF";
            string html = LinkEmail(shareUrl + "?xsdata=MDV8MDJ8&amp;sdata=abc", shareUrl);

            LinkResult result = LinkClassifier.ClassifyLink(html);

            Assert.Equal(EmailStatus.Delivered, result.Status);
            Assert.Equal(shareUrl, result.Link);
        }

        [Fact]
        public void RemovedLinkIsHeld()
        {
            string html = "<p>Please find your requested link for file: <b>test.doc</b>.</p><p>Link: .</p>";

            LinkResult result = LinkClassifier.ClassifyLink(html);

            Assert.Equal(EmailStatus.Held, result.Status);
            Assert.Equal("", result.Link);
        }

        [Fact]
        public void BareDelivrtoUrlWithoutAnchorIsRewritten()
        {
            // The URL is present in the body but no href opens it
            string html = $"<p>Please find your requested link for file: <b>test.doc</b>.</p><p>Link: {DelivrtoLink}.</p>";

            Assert.Equal(EmailStatus.Rewritten, LinkClassifier.ClassifyLink(html).Status);
        }

        [Fact]
        public void UppercaseAndSelfClosingWbrTagsAreStripped()
        {
            string html = LinkEmail(DelivrtoLink, DelivrtoLink).Replace("<wbr>", "<WBR />");

            Assert.Equal(EmailStatus.Delivered, LinkClassifier.ClassifyLink(html).Status);
        }

        [Theory]
        [InlineData(DelivrtoLink, true)]
        [InlineData(DelivrtoLink + "?a=b", true)]
        [InlineData(DelivrtoLink + "/", true)]
        [InlineData(DelivrtoLink + "#x", true)]
        [InlineData(DelivrtoLink + "0", false)]
        [InlineData("https://example.com/", false)]
        public void HrefPointsAtRequiresBoundaryAfterUrl(string href, bool expected)
        {
            Assert.Equal(expected, LinkClassifier.HrefPointsAt(href, DelivrtoLink));
        }
    }
}

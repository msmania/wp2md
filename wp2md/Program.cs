// Command to generate Microsoft.mshtml.dll
//
// tlbimp C:\Windows\system32\mshtml.tlb^
//        /namespace:Microsoft.mshtml^
//        /productversion:11.00.16299.15^
//        /asmversion:11.00.16299.15^
//        /out:Microsoft.mshtml.dll

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Microsoft.mshtml;

namespace wp2md {
  public class ImageInBlockQuote : Exception {
    public readonly string image_source_;
    public ImageInBlockQuote(string image_source) {
      image_source_ = image_source;
    }
  }

  public class WPXML {
    const string ELEM_STATUS = "{http://wordpress.org/export/1.2/}status";
    const string ELEM_TITLE = "title";
    const string ELEM_POSTDATE = "{http://wordpress.org/export/1.2/}post_date";
    const string ELEM_POSTDATE_GMT = "{http://wordpress.org/export/1.2/}post_date_gmt";
    const string ELEM_POSTTYPE = "{http://wordpress.org/export/1.2/}post_type";
    const string ELEM_POSTNAME = "{http://wordpress.org/export/1.2/}post_name";
    const string ELEM_DATE = "pubDate";
    const string ELEM_CATEGORY = "category";
    const string ELEM_CONTENT = "{http://purl.org/rss/1.0/modules/content/}encoded";

    static private void CollectCategories(IEnumerable<XElement> categoryElements,
                                   ref List<string> tags,
                                   ref List<string> categories) {
      foreach (var category in categoryElements) {
        var domain = category.Attribute("domain").Value;
        if (domain == "post_tag") {
          tags.Add(category.Value.ToString());
        }
        else if (domain == "category") {
          categories.Add(category.Value.ToString());
        }
        else {
          throw new Exception("Unknown category domain");
        }
      }
    }

    static private DateTimeOffset GetPostDate(XElement elemItem) {
      var local = elemItem.Elements(ELEM_POSTDATE).First();
      var gmt = elemItem.Elements(ELEM_POSTDATE_GMT).First();
      if (local != null && gmt != null) {
        var t_utc = DateTime.Parse(gmt.Value.ToString(),
                                   CultureInfo.CurrentCulture,
                                   DateTimeStyles.AssumeUniversal);
        var t_local = DateTime.Parse(local.Value.ToString(),
                                     CultureInfo.CurrentCulture,
                                     DateTimeStyles.AssumeUniversal);
        var offset = new DateTimeOffset(t_utc.ToUniversalTime());
        return offset.ToOffset(t_local - t_utc);
      }

      //var date = elemItem.Elements(ELEM_DATE).First();
      //if (date != null) {
      //  return DateTime.Parse(date.Value.ToString(),
      //                        CultureInfo.CurrentCulture,
      //                        DateTimeStyles.AssumeUniversal);
      //}

      throw new Exception("No date info");
    }

    private enum NodeType : int {
      element = 1,
      text = 3,
      comment = 8,
    }

    static private Dictionary<string, Func<WPXML, IHTMLDOMNode, string>> MDGenerators =
      new Dictionary<string, Func<WPXML, IHTMLDOMNode, string>>() {
      {"HTML", (p, n) => p.InnerText(n)},
      {"HEAD", (p, n) => ""},
      {"BODY", (p, n) => p.InnerText(n)},
      {"TITLE", (p, n) => ""},

      {"BR", (p, n) => "<br />\r\n"},
      {"A", (p, n) => p.Anchor(n)},
      {"IMG", (p, n) => p.Image(n)},

      {"P", (p, n) => "\r\n" + p.InnerText(n) + "\r\n\r\n"},
      {"DIV", (p, n) => p.Div(n)},

      {"B", (p, n) => "**" + p.InnerText(n) + "**"},
      {"STRONG", (p, n) => "**" + p.InnerText(n) + "**"},
      {"S", (p, n) => "~~" + p.InnerText(n) + "~~"},
      {"I", (p, n) => "_" + p.InnerText(n) + "_"},
      {"FONT", (p, n) => p.Font(n)},
      {"SPAN", (p, n) => p.Span(n)},
      
      {"UL", (p, n) => p.List(n, ordered:false)},
      {"OL", (p, n) => p.List(n, ordered:true)},
      {"LI", (p, n) => p.ListItem(n)},

      {"BLOCKQUOTE", (p, n) => p.BlockQuote(n)},

      {"H1", (p, n) => p.Heading(n, 1)},
      {"H2", (p, n) => p.Heading(n, 2)},
      {"H3", (p, n) => p.Heading(n, 3)},
      {"H4", (p, n) => p.Heading(n, 4)},
      {"H5", (p, n) => p.Heading(n, 5)},
      {"H6", (p, n) => p.Heading(n, 6)},
    };

    public string assetDirectory_ = "assets/";
    private string prefixOfCache_ = "cache_";
    private string currentPost_;

    private Stack<bool> listModesOrdered_ = new Stack<bool>();
    private bool containsImage_ = false;
    private bool processingBlockquote_ = false;
    private string processingHyperlink_ = "";

    private bool ignoreTextnodes_ = false;
    private bool dontSanitizeText_ = false;

    public bool skipDownload_ = true;
    public bool handleImgInBlockQuoteException_ = false;

    // Returns the name of the cached file
    private string DownloadFile(string uri) {
      var filename = Path.GetFileName(uri);
      var questionmark = filename.IndexOf('?');
      if (questionmark != -1) {
        filename = filename.Substring(0, questionmark);
      }
      filename = prefixOfCache_ + filename;
      var client = new System.Net.WebClient();
      if (!skipDownload_) {
        client.DownloadFile(uri, Path.Combine(assetDirectory_, filename));
        System.Diagnostics.Debug.WriteLine("Download: {0}\t{1}", uri, filename);
      }
      return Uri.EscapeDataString(filename);
    }

    static private string GetAttribute(IHTMLDOMNode node, string name) {
      var attrs = node.attributes as IHTMLAttributeCollection2;
      var attr = attrs.getNamedItem(name);
      return attr != null && attr.specified ? attr.nodeValue.ToString() : "";
    }

    static IHTMLStyle GetStyle(IHTMLDOMNode node) {
      var elem = node as IHTMLElement;
      return elem?.style;
    }

    private string Heading(IHTMLDOMNode node, UInt32 size) {
      var elem = node as IHTMLElement;
      return string.Concat("\r\n",
                           new string('#', (int)Math.Min(size, 6u)),
                           " ",
                           elem.innerHTML,
                           "\r\n\r\n");
    }

    private string Font(IHTMLDOMNode node) {
      var color = GetAttribute(node, "COLOR");
      return color.Length > 0 ? OuterText(node)
                              : InnerText(node);
    }

    private string Span(IHTMLDOMNode node) {
      var style = GetStyle(node);
      return style != null && style.color != null
             ? OuterText(node)
             : InnerText(node);
    }

    private string Image(IHTMLDOMNode node) {
      containsImage_ = true;
      var source = processingHyperlink_.Length > 0
                   ? processingHyperlink_
                   : GetAttribute(node, "SRC");

      if (processingBlockquote_) {
        if (handleImgInBlockQuoteException_) {
          // Image in Blockquote is omitted due to innerText.
          // Need to take it out from Blockquote manually.
          Console.WriteLine("> " + source + " in BlockQuote in " + currentPost_);
        }
        else
          throw new ImageInBlockQuote(source);
      }

      return assetDirectory_ == null
             ? string.Format("![]({0})", source)
             : string.Format("![]({{{{site.assets_url}}}}{0})",
                             DownloadFile(source));
    }
    
    private string Anchor(IHTMLDOMNode node) {
      containsImage_ = false;
      processingHyperlink_ = GetAttribute(node, "HREF");

      dontSanitizeText_ = true;
      var childText = InnerText(node);
      dontSanitizeText_ = false;

      string output = containsImage_
                      ? childText
                      : string.Format("[{0}]({1})",
                                      childText,
                                      GetAttribute(node, "HREF"));

      processingHyperlink_ = "";
      containsImage_ = false;
      return output;
    }

    private string Div(IHTMLDOMNode node) {
      var attr_class = GetAttribute(node, "class");
      var content = InnerText(node);
      return content.StartsWith("Livedoor \u30BF\u30B0")
             ? ""
             : InnerText(node);
    }

    private string BlockQuote(IHTMLDOMNode node) {
      var childText = "";
      processingBlockquote_ = true;
      childText = InnerText(node);
      processingBlockquote_ = false;

      childText = "";
      if (node is IHTMLElement element) {
        childText = element.innerText;
      }
      return "\r\n```\r\n" + childText + "\r\n```\r\n";
    }

    private string List(IHTMLDOMNode node, bool ordered) {
      listModesOrdered_.Push(ordered);
      ignoreTextnodes_ = true;
      var childText = InnerText(node);
      ignoreTextnodes_ = false;
      listModesOrdered_.Pop();
      return "\r\n" + childText + "\r\n";
    }

    private string ListItem(IHTMLDOMNode node) {
      var ignoreTextnodes_old = ignoreTextnodes_;
      ignoreTextnodes_ = false;
      var childText = InnerText(node);
      ignoreTextnodes_ = ignoreTextnodes_old;
      return string.Concat(listModesOrdered_.Peek() ? "1. " : "- ",
                           childText,
                           "\r\n");
    }

    private string InnerText(IHTMLDOMNode node) {
      string childText = "";
      var child = node.firstChild as IHTMLDOMNode;
      while (child != null) {
        childText += NodeToMD(child);
        child = child.nextSibling;
      }
      return childText;
    }

    private string OuterText(IHTMLDOMNode node) {
      if (node is IHTMLElement element) {
        return element.outerHTML;
      }
      return node.ToString();
    }

    static private string SanitizeText(string original) {
      var sanitized = original;

      var whitespaces = new System.Text.RegularExpressions.Regex(@"\s+");
      sanitized = whitespaces.Replace(sanitized, " ");

      sanitized = sanitized.Replace("[", "&#x5b;");
      sanitized = sanitized.Replace("]", "&#x5d;");
      sanitized = sanitized.Replace("<", "&lt;");
      sanitized = sanitized.Replace(">", "&gt;");

      var underscore = new System.Text.RegularExpressions.Regex(@"\b_\S*_");
      sanitized = underscore.Replace(sanitized, match => {
        return "\\" + match.Value;
      });

      return sanitized;
    }

    private string NodeToMD(IHTMLDOMNode rootNode) {
      var type = (NodeType)rootNode.nodeType;
      if (type == NodeType.comment)
        return "";
      else if (type == NodeType.text) {
        if (ignoreTextnodes_)
          return "";

        var rawText = rootNode.nodeValue.ToString();
        if (rawText == "\n")
          return "";

        return dontSanitizeText_ ? rawText : SanitizeText(rawText);
      }

      if (rootNode is IHTMLElement elem
          && (rootNode as IHTMLBRElement) == null
          && (rootNode as IHTMLImgElement) == null
          && elem.innerHTML == null) {
        return "";
      }

      Func<WPXML, IHTMLDOMNode, string> func = (p, n) => p.OuterText(n);
      try {
        func = MDGenerators[rootNode.nodeName];
      }
      catch (KeyNotFoundException) {
        System.Diagnostics.Debug.WriteLine("Unhandled element: " + rootNode.nodeName);
      }
      return func(this, rootNode);
    }

    public string HTMLToMD(string htmltext) {
      var doc = new HTMLDocument();
      var doc2 = doc as IHTMLDocument2;
      doc2.write(@"<!DOCTYPE html>
<html>
<head>
<meta http-equiv=""x-ua-compatible"" content=""IE=edge"">
</head>
<body>");
      doc2.write(htmltext);
      doc2.write(@"</body></html>");
      doc2.close();
      var root = (doc as IHTMLDocument3).documentElement;
      var full = (root as IHTMLElement).outerHTML;
      return NodeToMD(root as IHTMLDOMNode);
    }

    static string GenerateHeader(string title,
                                 DateTimeOffset date,
                                 List<string> categories,
                                 List<string> tags) {
      var headerLines = new List<string>() {
        "---",
        "layout: post",
        "title: \"" + title + "\"",
        "date: " + date.ToString("yyyy-MM-dd HH:mm:ss.fff zzz")
      };

      if (categories.Count > 0) {
        headerLines.Add("categories:");
        foreach (var category in categories) {
          headerLines.Add("- " + category);
        }
      }
      if (tags.Count > 0) {
        headerLines.Add("tags:");
        foreach (var tag in tags) {
          headerLines.Add("- " + tag);
        }
      }
      headerLines.Add("---\r\n");

      return string.Join("\r\n", headerLines);
    }

    public void Dump(string filepath, string output_path) {
      var doc = XDocument.Load(filepath);
      var items = doc.Descendants("item");
      foreach (var item in items) {
        var status = item.Elements(ELEM_STATUS).First();
        var type = item.Elements(ELEM_POSTTYPE).First();
        if (status != null
            && status.Value == "publish"
            && type.Value == "post") {
          var title = item.Elements(ELEM_TITLE).First().Value;
          var postname = item.Elements(ELEM_POSTNAME).First().Value;
          var date = GetPostDate(item);
          var tags = new List<string>();
          var categories = new List<string>();
          CollectCategories(item.Elements(ELEM_CATEGORY),
                            ref tags,
                            ref categories);

          prefixOfCache_ = date.ToString("yyyy-MM-dd-");
          currentPost_ = prefixOfCache_ + postname;
          Console.WriteLine("Processing .. " + currentPost_);

          var elem_content = item.Elements(ELEM_CONTENT).First();
          if (elem_content != null) {
            var output_fullpath = Path.Combine(output_path, currentPost_);
            var content_html = elem_content.Value;
            var content_md = HTMLToMD(content_html);
            var header = GenerateHeader(title, date, categories, tags);
            File.WriteAllText(output_fullpath + ".html",
                              content_html);
            File.WriteAllText(output_fullpath + ".md",
                              header + content_md);
          }
        }
      }
    }
  }

  class Program {
    static void Main(string[] args) {
      var wp = new WPXML {
        skipDownload_ = true,
        handleImgInBlockQuoteException_ = true,
        assetDirectory_ = @"D:\blog\assets"
      };
      wp.Dump(@"D:\blog\wordpress.2017-11-25.001.xml",
              @"D:\blog\output");
      wp.Dump(@"D:\blog\wordpress.2017-11-25.002.xml",
              @"D:\blog\output");
    }
  }
}

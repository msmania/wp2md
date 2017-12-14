using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace wp2md {
  [TestClass]
  public class Test_WPXML {
    WPXML wp;

    public Test_WPXML() {
      this.wp = new WPXML();
      wp.skipDownload_ = true;
    }

    public void Equal(string input, string expectedOutput) {
      var actual = wp.HTMLToMD(input);
      Assert.AreEqual(expectedOutput, actual);
    }

    [TestMethod]
    public void Single() {
      var longUrl = @"https://www.google.co.jp/search?num=100&safe=off&dcr=0&source=hp&ei=MHYoWrvQHYHd0ASt_YXICg&q=alphazero&oq=alphazero&gs_l=psy-ab.3..0i10k1j0j0i10k1l6j0i30k1l2.201.1696.0.1885.10.8.0.1.1.0.127.744.5j3.8.0....0...1c.1.64.psy-ab..1.9.746.0...0.xxMaIybw4Uw";
      Equal(@"<a href=""#123"">link</a>", @"[link](about:blank#123)");
      Equal(@"<a href=""" + longUrl + @""">link</a>",
            @"[link](" + longUrl + @")");
      Equal(@"<b>abc</b>", @"**abc**");
      Equal(@"<blockquote>&#160; // comment1<br />&#160; // comment2<br /></blockquote>",
            @"
```
  // comment1
  // comment2
```
");
      Equal(@"<i>abc</i>", @"_abc_");
      Equal(@"<img src=""lena.jpg""></img>", @"![]({{site.assets_url}}cache_lena.jpg)");
      Equal(@"<img src=""" + longUrl + @""" />", @"![]({{site.assets_url}}cache_search)");
      Equal(@"<p>hello</p>", @"
hello

");
      Equal(@"<p>hello<br />world</p>", @"
hello<br />
world

");
      Equal(@"<s>abc</s>", @"~~abc~~");
      Equal(@"<strong>abc</strong>", @"**abc**");

      Equal(@"<strong></strong>", "");
      Equal(@"<button>abc</button>", @"<button>abc</button>");
    }

    [TestMethod]
    public void Multi() {
      Equal(@"<p>1</p>
<p>2</p>", @"
1


2

");
      Equal(@"<strong>asterisks and <i>underscores</i></strong>",
            @"**asterisks and _underscores_**");
      Equal(@"<h1>header1</h1>
<h2>header2</h2>
<h3>header3</h3>
<h4>header4</h4>
<h5>header5</h5>
<h6>header6</h6>", @"
# header1


## header2


### header3


#### header4


##### header5


###### header6

");
    }

    [TestMethod]
    public void ImageUrl() {
      Equal(@"<a href=""http://host/image.bmp?abc""><img src=""lena.jpg"" ></a>",
            @"![]({{site.assets_url}}cache_image.bmp)");
    }

    [TestMethod]
    [ExpectedException(typeof(Exception),
      "Found IMG in Blockquote.  Need to resolve it manually.")]
    public void ImageInBlockQuote() {
      Equal(@"<blockquote>aaa<br />
<a href=""#""><img src=""lena.jpg"" /></a>
</blockquote>", @"
```
aaa
```
");
    }

    [TestMethod]
    public void LiveTag() {
      Equal(@"<div style=""margin: 0px; padding: 0px; float: none; display: inline;"">Livedoor タグ: 
<a href=""http://clip.livedoor.com/tag/VMware"" rel=""tag"">VMware</a>,
<a href=""http://clip.livedoor.com/tag/Windows"" rel=""tag"">Windows</a></div>",
      @"");
    }

    [TestMethod]
    public void Font() {
      Equal(@"<font>no attributes</font><br />
<font color=""#ff0000"">colored</font>",
            @"no attributes<br />
<font color=""#ff0000"">colored</font>");
      Equal(@"<span>no attributes</span><br />
<span style=""font-family:'Courier New'"">font</span><br />
<span style=""font-family:'Courier New';color:#ff0000"">colored</span>",
            @"no attributes<br />
font<br />
<span style='color: rgb(255, 0, 0); font-family: ""Courier New"";'>colored</span>");
    }

    [TestMethod]
    public void List() {
      Equal(@"<ul>
<li>apple</li>a
<li>google</li>b
<li>microsoft</li>c
</ul>", @"
- apple
- google
- microsoft

");
      Equal(@"<ol>
<li>apple</li>
<li>google</li>
<li>microsoft</li>
</ol>", @"
1. apple
1. google
1. microsoft

");
      Equal(@"<ol><li>item1</li><li>item2<br />
<img src=""lena.jpg"" /></li></ol>
<p>paragraph</p>", @"
1. item1
1. item2<br />
![]({{site.assets_url}}cache_lena.jpg)


paragraph

");
    }

    [TestMethod]
    public void Text() {
      Equal(@"<p><> [def]</p>", @"
&lt;&gt; &#x5b;def&#x5d;

");
      Equal(@"<p>_abc_ __def__ 1_abc_2</p>", @"
\_abc_ \__def__ 1_abc_2

");
      Equal(@"<p><a href=""http://a/_b_/c"">http://a/_b_/c</a></p>", @"
[http://a/_b_/c](http://a/_b_/c)

");
      Equal(@"<p>a<br />    <br /> 
 </p>", @"
a<br />
 <br />
 

");
    }
  }
}

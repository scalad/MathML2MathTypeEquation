## MathML2MathTypeEquation ##


### Support ###
* MathML公式类使用MathType公式进行转换
* HTML格式导出Word(支持表格、图片、排序标签、转义字符等)

### Required ###
* Office2003+(Word And Excel) Best of 2007
* .Net Frameword4.0+
* MathType6.9+[关于MathType6.9破解](http://download.csdn.net/detail/qq_20545159/9921565)

### Introduce ###

这篇文章主要介绍怎么使用[MathType](https://www.dessci.com/en/reference/sdk/)把[MathML](https://zh.wikipedia.org/wiki/%E6%95%B0%E5%AD%A6%E7%BD%AE%E6%A0%87%E8%AF%AD%E8%A8%80)转换MathType类型的公式对象并嵌入到Word中。

最近，这里有个需求是需要转换一大批的MathML文件到Word文档中，如果使用[Open-XML-SDK](https://github.com/OfficeDev/Open-XML-SDK)是非常容易实现的，你可以参考[https://github.com/scalad/MathML2Word](https://github.com/scalad/MathML2Word)，但是，最重要的是这不是想要的结果，因为经过Open-XML-SDK的转换，这个公式的类型变成了[OMML(Office Math Markup Language)](https://en.wikipedia.org/wiki/Mathematical_markup_language)格式的，什么是OMML呢？

我们知道，微软的Word包含了公式编辑器，其实它是一个缩小版本的MathType，这个从上世纪word出现时已经开始了。直到2007年，word才允许使用[图形用户界面](https://en.wikipedia.org/wiki/Graphical_user_interface)输入公式，并且转换为像MATHML格式的标记语言。随着微软发布了[Microsoft Office 2007](https://en.wikipedia.org/wiki/Microsoft_Office_2007) 和[Office Open XML file formats](https://en.wikipedia.org/wiki/Office_Open_XML_file_formats),微软引进了一个使用新的格式的公式编辑器，即所谓的`Office Math Markup Language(OMML)`，OMML与原来的公式编辑器存在着兼容性问题，因此很多学术官网都拒绝使用Microsoft Office写的文档。

Mathtype公式编辑器是基于宏或是VB编出来的，实际上，在Office2007之前的版本中，微软一直使用的是MathType提供的缩小版本的MathType公式编辑器，想要使用完整公式编辑器的还需要用户到MathType去买(没错，在长达15年的时间里，所有Office都自带MathType的缩小版)，直到2007之后，微软才开发出属于自己的一套公式编辑器，它的公式类型是OMML(Office Math Markup Language)，并且和原有的MathType公式类型不兼容，因此，有许多学术网站都明确提出了不使用Office2007以及后面的版本。

MathType SDK是针对MathType工具用VB完成的一套开发工具包，它允许开发人员改造、扩展、修改或者创建命令等，并且官方文档中提供了.NET平台上SDK的实现，你可以很方便的使用C#调用它们。如下图是.NET平台上公式支持的输入输出的格式：

![](https://github.com/scalad/MathML2MathTypeEquation/blob/master/doc/image/MTSDKDN.png)

EquationInput(公式输入)、EquationOutput(公式输出)和MTSDK(MathType连接、释放)作为ConverttEquation的成员变量，ConverttEquation初始化时首先完成了MTSDK对象的初始化。MTSDK包含了两个方法，Init()和DeInit()，用来连接MathType服务和释放服务。而后调用ConvertEquation中的Convert方法完成它们两个所支持的文件格式的转换。

目前采用的方式是使用EquationInputFileText类从磁盘文件中读入MathML数据类型的数据，然后使用EquationOutputClipboardText输出到系统的剪切板中，从剪切板中获取到该公式的对象并写入到Word文档中，当文件读取并转换完成后，生成Wrod文档并保存。

### Question ###
1、特殊公式映射转换

```C#
static MathML()
{
    equations.Add("∆", "&#x0394;");
    equations.Add("<mo><</mo>", "<mo>&#x003C;</mo>");
    equations.Add("<mo>></mo>", "<mo>&#x003E;</mo>");
    equations.Add("<mo>⪈</mo>", "<mo>&#x2269;</mo>");
    equations.Add("<mo>⪇</mo>", "<mo>&#x2268;</mo>");
    equations.Add("<mo>≢</mo>", "<mo>&#x2260;</mo>");
    equations.Add("<mo>⊝</mo>", "<mo>&#x2296;</mo>");
    equations.Add("<mo>·</mo>", "<mo>&#x22C5;</mo>");
    equations.Add("<mo>⋅</mo>",  "<mo>&#x22C5;</mo>");
    equations.Add("<mo>&nbsp;</mo>", "");
    equations.Add("<mtext>&#x00A1;&#x00CE;</mtext>", "<mo>&#x2225;</mo>");//||
}
```
由于原来的数据是经过[MathJax](https://github.com/mathjax/MathJax)转换而来，因此，该平台上有的图标或者是公式在MathType上是无法转换，也就是说MathType无法识别这些图标，所以我们需要把这些不能识别的公式进行映射替换，替换为MathType可以识别的Unicode编码.

2、关于MathType弹出缺少字体的对话框
因为在上面的字符映射上可能使用了不同字体的Unicode编码，很有可能导致这个问题：
MathType需要一个新版的MT Extra（TrueType）字体，请重新安装MathType（用CD或下载安装）以便获取正确的字体

这时，一些符号就无法使用。下面介绍解决以上问题的办法。

1，进入文件夹X:\Windows\Fonts（X为系统盘盘符，一般是C），如果有MT Extra(TrueType)字体或者快捷方式，则将其删除。

2，进入MathType安装目录的\Fonts\TrueType\文件夹，将里面的mtextra.ttf字体复制至文件夹X:\Windows\Fonts（粘贴时可能会有安装字体的提示）。

### Operation ###
![](https://github.com/scalad/MathML2MathTypeEquation/blob/master/doc/effect.gif)
## MathML2MathTypeEquation ##

这篇文章主要介绍怎么使用[MathType](https://www.dessci.com/en/reference/sdk/)把[MathML](https://zh.wikipedia.org/wiki/%E6%95%B0%E5%AD%A6%E7%BD%AE%E6%A0%87%E8%AF%AD%E8%A8%80)转换MathType类型的公式对象并嵌入到Word中。

最近，这里有个需求是需要转换一大批的MathML文件到Word文档中，如果使用[Open-XML-SDK](https://github.com/OfficeDev/Open-XML-SDK)是非常容易实现的，你可以参考[https://github.com/scalad/MathML2Word](https://github.com/scalad/MathML2Word)，但是，最重要的是这不是想要的结果，因为经过Open-XML-SDK的转换，这个公式的类型变成了[OMML(Office Math Markup Language)](https://en.wikipedia.org/wiki/Mathematical_markup_language)格式的，什么是OMML呢？

我们知道，微软的Word包含了公式编辑器，其实它是一个缩小版本的MathType，这个从上世纪word出现时已经开始了。直到2007年，word才允许使用[图形用户界面](https://en.wikipedia.org/wiki/Graphical_user_interface)输入公式，并且转换为

As we know,Microsoft Word included Equation Editor, a limited version of MathType, until 2007.These allow entering formulae using a [graphical user interface](https://en.wikipedia.org/wiki/Graphical_user_interface), and converting to standard markup languages such as MathML.With Microsoft's release of [Microsoft Office 2007](https://en.wikipedia.org/wiki/Microsoft_Office_2007) and the [Office Open XML file formats](https://en.wikipedia.org/wiki/Office_Open_XML_file_formats), they introduced a new equation editor which uses a new format, "Office Math Markup Language" (OMML). The lack of compatibility led some prestigious scientific journals to refuse to accept manuscripts which had been produced using Microsoft Office 2007.

[关于MathType6.9破解](http://download.csdn.net/detail/qq_20545159/9921565)
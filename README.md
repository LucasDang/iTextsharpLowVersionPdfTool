## 关于在.net Core使用iTextsharp导出pdf

**前言：这两个星期一直在弄在Core上导出pdf文件，并且pdf文件是属于比较复杂，所以需要用到模板来填充数据，遇到了很多的问题！**

 - 使用的是Core,iTextsharp最新的版本5.x与之不兼容，并且最新的版本已经不开源了，商业用途需要收费的。
 - 我使用的是版本4.x，还是零几年的一个版本。这个版本有一个问题就是凡是编码的字体都会嵌入到pdf中，所以一个原本只有几十k大小的pdf变成了4m多，而且这只是一个pdf，如果有批量导出的话，五十个就是几百m了，可怕！
 - 最关键的是所有的注释全部都是英文，而且是一段落一段落的，而且相关的问题百度到的压根没有参考价值，只能去stackoverflow找答案，看的脑子都浆糊了...
 
所以在此将我所遇到的一些坑记录一下，防止有人重走我的老路！

---

###详细步骤


**pdf模板制作**

  下载adobe acrobat pro这个软件，当然是收费么，不过我大天朝嘛（ps:大家有钱的还是要支持支持正版呀，当然我很穷。）然后打开一个pdf文档，点击导航栏的表单，可以看到有很多选项：表单向导是自动帮你选择区域（将空白的部分划分成一个区域），你也可以自己添加区域来编辑。类似下图，那一个一个蓝色的就是一个一个区域，回头可以直接赋值的

![模板区域制作](https://ooo.0o0.ooo/2017/06/21/594a1c6fbd422.png)

---

**pdf模板赋值**

 - 首先需要导入iTextsharp这个库，如果你不是Core的话直接导入最新版本就好了，如果你用Core的话我在里面貌似没找到5.x之前的版本，我导的是[iTextSharp.LGPLv2.Core](https://github.com/VahidN/iTextSharp.LGPLv2.Core),这是伊朗的一位C#开发者封装的一个工具类。
 - 导入库之后就可以来对模板进行赋值,代码如下

		//获取中文字体，第三个参数表示为是否潜入字体，但只要是编码字体就都会嵌入。
		BaseFont baseFont = BaseFont.CreateFont(@"C:\Windows\Fonts\simsun.ttc,1",BaseFont.IDENTITY_H,BaseFont.NOT_EMBEDDED);
		//读取模板文件
        PdfReader reader = new PdfReader(@"C:\Users\Administrator\Desktop\template.pdf");

        //创建文件流用来保存填充模板后的文件
        MemoryStream stream = new MemoryStream();
                
        PdfStamper stamp = new PdfStamper(reader, stream);
        //设置表单字体，在高版本有用，高版本加入这句话就不会插入字体，低版本无用
        //stamp.AcroFields.AddSubstitutionFont(baseFont);

        AcroFields form = stamp.AcroFields;

        //表单文本框是否锁定
        stamp.FormFlattening = true;

        //填充表单,para为表单的一个（属性-值）字典
        foreach (KeyValuePair<string, string> parameter in para)
        {
            //要输入中文就要设置域的字体;
            form.SetFieldProperty(parameter.Key, "textfont", baseFont, null);
            //为需要赋值的域设置值;
             form.SetField(parameter.Key, parameter.Value);
         }
	
	 还可以对模板添加图片，代码如下

		
        // 通过域名获取所在页和坐标，左下角为起点
        float[] positions = form.GetFieldPositions("sender");
        int pageNo = (int)positions[0];
        float x = positions[1];
        float y = positions[2];
        // 读图片
        Image image = Image.GetInstance(@"C:\Users\Administrator\Desktop\02.png");
        // 获取操作的页面
        PdfContentByte under = stamp.GetOverContent(pageNo);
        // 根据域的大小缩放图片
        image.ScaleToFit(positions[3] - positions[1], positions[4] - positions[2]);
        // 添加图片
        image.SetAbsolutePosition(x, y);
        under.AddImage(image);
		
      这里大家可能对**positions**不太了解，他的注释：
 
	>Gets the field box positions in the document. The return >is an array of float multiple of 5. For each of this groups the values are: [page, llx, lly, urx, ury]. The coordinates have the page rotation in consideration.
	

	大概意思就是获取pdf一个域的坐标，返回了一个数组，盛放的分别是：[当前是第几页]，[区域左下方的X坐标]，[区域左下方的Y坐标]，[区域右上方的X坐标]，[区域右上方的Y坐标]。通过四个点就知道这块区域的大小，然后再定位到是第几页。（语文水平不太高，听不懂就.....see more times）

	

	最后按顺序关闭io流
	
		stamp.Close();
        reader.Close();


**这里是单个模板赋值，单个模板生成多个pdf就加个for循环就好，多个模板的话也差不多，单个模板搞定了，其他都简单了**

--- 



**合并生成的多个pdf文件**

  在上面我这里没有生成pdf文件，而是用MemoryStream将赋值完成的pdf放到了内存中，然后再将他们合并发送到客户端直接下载，这样就避免了许多的文件io操作，提升点效率。（当然如果文件太多太大的话可能会内存溢出而挂掉吧）

代码如下：

	Document document = new Document();

    //将合并后的pdf依旧以流的方式保存在内存中
    MemoryStream mergePdfStream = new MemoryStream();

    //这里用的是smartCopy，整篇文档只会导入一份字体。属于可接受范围内
    PdfSmartCopy copy = new PdfSmartCopy(document, mergePdfStream);

    document.Open();

    for (int i = 0; i < pdfStreams.Count; i++)
    {
	   for循环读取之前生成的单个pdf字节数组，这里将流保存在数组中传递过来使用的话不行，会报错，我也不知道为什么，所以用了字节数组（如果有人知道的话还请告诉我，不胜感激）
       byte[] stream = pdfStreams[i];
       PdfReader reader = new PdfReader(stream);

       //for循环新增文档页数，并copy pdf数据
       document.NewPage();
       PdfImportedPage imported = copy.GetImportedPage(reader, 1);
       copy.AddPage(imported);

        reader.Close();
     }
     copy.Close();
     document.Close();

     //这里将合并后的pdf流传出去，方法我没全复制过来。（你们也不能光复制就想着直接用吧，好歹动下头脑）
     return mergePdfStream;

---

**pdf在浏览器中下载**

这是最后一步了，我用的是MVC,不过不管用什么都一样的原理

	FileResult fileResult = new FileContentResult(mergePdfStream.ToArray(), "application/pdf");
    fileResult.FileDownloadName = "4.pdf";

然后直接返回fileResult就行，因为它是继承于ActionResult的，应该是MVC框架的一个专门返回文件的类。（.net就是爽，微软封装的太好了）

---

**and最终**


趁还在，多陪陪身边的人，莫等到失去才开始后悔。

另外厦门是个好地方！

---

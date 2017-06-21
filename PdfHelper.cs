using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;

namespace ExpressSystem.Helpers
{
    public class PdfHelper
    {

        const string ResourcesFolder = "pdfResource";

        //获取根目录
        public static string GetBaseDir()
        {
            var currentAssembly = typeof(PdfHelper).GetTypeInfo().Assembly;
            var root = Path.GetDirectoryName(currentAssembly.Location);
            var idx = root.IndexOf($"{Path.DirectorySeparatorChar}bin", StringComparison.OrdinalIgnoreCase);
            return root.Substring(0, idx);
        }
        //获取模板路径
        public static string GetPdfTemplatePath()
        {
            return Path.Combine(GetBaseDir(), ResourcesFolder, "pdfTemplate", "template.pdf");
        }



        /// <summary>
        /// 根据数据填充模板并获取一个一个pdf文件流
        /// </summary>
        /// <param name="listPara">数据参数</param>
        /// <returns>所有的pdf文件字节数组，并装在一个数组中</returns>
        public static List<byte[]> GetPdfs(List<Dictionary<string, string>> listPara)
        {
            //获取中文字体，第三个参数表示为是否潜入字体，但只要是编码字体就都会嵌入。
            BaseFont baseFont = BaseFont.CreateFont(@"C:\Windows\Fonts\simsun.ttc,1", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
           
            
            List<byte[]> pdfStreams = new List<byte[]>();

            foreach(Dictionary<string,string> para in listPara)
            {
                //读取模板文件
                PdfReader reader = new PdfReader(@"C:\Users\Administrator\Desktop\template.pdf");

                //创建文件流用来保存填充模板后的文件
                MemoryStream stream = new MemoryStream();
                
                PdfStamper stamp = new PdfStamper(reader, stream);

                stamp.AcroFields.AddSubstitutionFont(baseFont);

                AcroFields form = stamp.AcroFields;
                stamp.FormFlattening = true;//表单文本框锁定
                //填充表单
                foreach (KeyValuePair<string, string> parameter in para)
                {
                    //要输入中文就要设置域的字体;
                    form.SetFieldProperty(parameter.Key, "textfont", baseFont, null);
                    //为需要赋值的域设置值;
                    form.SetField(parameter.Key, parameter.Value);
                }

                //添加图片
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
                
               
                stamp.Close();
                reader.Close();

                byte[] result = stream.ToArray();
                pdfStreams.Add(result);
            }
            return pdfStreams;
        }

        

        /// <summary>
        /// 合并生成的pdf文件流
        /// </summary>
        /// <param name="pdfStreams">多个pdf文件字节</param>
        /// <returns>返回合并之后的pdf文件流</returns>
        public static MemoryStream MergePdfs(List<byte[]> pdfStreams)
        {

            try
            {
                Document document = new Document();
                MemoryStream mergePdfStream = new MemoryStream();
                //这里用的是smartCopy，整篇文档只会导入一份字体。属于可接受范围内
                PdfSmartCopy copy = new PdfSmartCopy(document, mergePdfStream);

                document.Open();
                for (int i = 0; i < pdfStreams.Count; i++)
                {
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

                return mergePdfStream;
            }
            catch (Exception ex)
            {
                ex.ToString();
                return null;
            }


        }


    }
}

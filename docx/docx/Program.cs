using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Collections;
using Microsoft.Office.Interop.Word;
using System.IO;

namespace docx
{
    class Program
    {
        public static void Main()
        {

            
                Random ran = new Random();
                string Date = DateTime.Now.ToLongDateString().ToString();
                string path = @"E:\Job\54\docx\report\" + Date; //生成的目录
                
                string TimeNow = DateTime.Now.Hour.ToString() + DateTime.Now.Minute.ToString() + DateTime.Now.Second.ToString() + @"-" + ran.Next(100, 1000);
                string SavedPath = path + @"\" + TimeNow;
                string templatePath = @"E:\Job\54\docx\templete\profile1.doc";
                string ImgPath = @"E:\Job\54\docx\img/1.png";
            try
                {
                    if (!Directory.Exists(path))
                    {
                        Directory.CreateDirectory(path);
                    }

                }
                catch { }
                Report report = new Report();
                report.CreateNewDocument(templatePath); //模板路径
                report.InsertValue("name", "赵晓乐");//在书签“name”处插入值
                report.InsertValue("xingming", "赵晓乐");//在书签“name”处插入值
                report.InsertValue("age", "26");//
                report.InsertValue("sex", "男");//
                report.InsertValue("minzu", "汉");//
                report.InsertValue("jiguan", "张家口张北");//
                report.InsertValue("chushengriqi", "19930414");//
                report.InsertValue("zhengzhimianmao", "党员");//
                report.InsertValue("xueli", "研究生");//
                report.InsertValue("zhuanye", "测试计量技术");//
                report.InsertValue("tongxundizhi", "中国*******大学");//
                report.InsertValue("email", "10000*****@qq.com");//
                report.InsertValue("youbian", "******");//
                report.InsertValue("tel", "135********");//
                string text = "长期从事电脑操作者，应多吃一些新鲜的蔬菜和水果，同时增加维生素A、B1、C、E的摄入。为预防角膜干燥、眼干涩、视力下降、甚至出现夜盲等，电 脑操作者应多吃富含维生素A的食物，如豆制品、鱼、牛奶、核桃、青菜、大白菜、空心菜、西红柿及新鲜水果等。";
                report.InsertText("jiaoyubeijing", text);
                report.InsertText("gerenjineng", text);
                report.InsertText("huojiangqingkuang", text);
                
                report.InsertPicture("img", ImgPath, 75 , 115); //书签位置，图片路径，图片宽度，图片高度


            Table table = report.InsertTable("table", 2, 3, 0); //在书签“Bookmark_table”处插入2行3列行宽最大的表





            report.SaveDocument(SavedPath); //文档路径
           

           
        }
    }

}

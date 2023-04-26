using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
namespace 串口助手
{
    class Test
    {
        //private const string FILE_NAME2 = "数据库连接.txt";
       // private const string FILE_NAME = "数据库连接.txt";
        public void WriteText(string FILE_NAME ,string str)
        {
            //FileStream fs = new FileStream(FILE_NAME, FileMode.Create);
            //BinaryWriter w = new BinaryWriter(fs);
            StreamWriter sw = new StreamWriter(FILE_NAME,false, System.Text.Encoding.UTF8);
            sw.Write(str);
            sw.Close();
            //fs.Close();
        }
        public string ReadText(string FILE_NAME)
        {
            string input, str;
            using (StreamReader sr = File.OpenText(FILE_NAME))
            {
                while ((input = sr.ReadLine()) != null)
                {
                    str = input.ToString();
                    sr.Close();
                    return str;
                }
                return "";
            }
        }
        public string[] ReadText2(string FILE_NAME)
        {
            string input, str;
            string[] str2 = {"",""};
            using (StreamReader sr = File.OpenText(FILE_NAME))
            {
                
               
                    if((input = sr.ReadLine()) != null )
                    {
                        str = input.ToString();
                        str2[0] = str;
                        Console.WriteLine(str[0]+"测试2");
                        if ((input = sr.ReadLine()) != null  )
                        {
                            str = input.ToString();
                            str2[1] = str;
                            Console.WriteLine(str[1] + "测试3");
                        }
                    }
                    sr.Close();
                    return str2;
             
                
               
            }
        }
    }
}

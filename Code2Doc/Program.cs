using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Code2Doc
{
    class Program
    {
        static List<string> FindFile2(string sSourcePath)
        {
            List<String> list = new List<string>();
            //遍历文件夹
            DirectoryInfo theFolder = new DirectoryInfo(sSourcePath);
            FileInfo[] thefileInfo = theFolder.GetFiles("*.*", SearchOption.TopDirectoryOnly);
            foreach (FileInfo NextFile in thefileInfo)  //遍历文件
                list.Add(NextFile.FullName);
            //遍历子文件夹
            DirectoryInfo[] dirInfo = theFolder.GetDirectories();
            foreach (DirectoryInfo NextFolder in dirInfo)
            {
                //list.Add(NextFolder.ToString());
                FileInfo[] fileInfo = NextFolder.GetFiles("*.*", SearchOption.AllDirectories);
                foreach (FileInfo NextFile in fileInfo)  //遍历文件
                    list.Add(NextFile.FullName);
            }
            return list;
        }

        static string getSrcPath(string[] args)
        {
            var srcPath = String.Empty;
            if (args.Length > 0)
            {
                srcPath = args[0];
            }

            if (Directory.Exists(srcPath) == false)
            {
                Console.WriteLine("src path not exists");
                Console.WriteLine("pls type or paste your src path:");

                return getSrcPath(new string[] { Console.ReadLine() });
            }

            return srcPath;
        }

        static void Main(string[] args)
        {
            string dir = AppDomain.CurrentDomain.BaseDirectory;
            string TemPath = System.IO.Path.Combine(dir, "test.docx");

            var srcPath = getSrcPath(args);

            var files = FindFile2(srcPath);

            foreach (var file in files)
            {
                Console.WriteLine(file);

                var fileName = System.IO.Path.GetFileName(file);

                Console.WriteLine(fileName);

                var fileType = System.IO.Path.GetExtension(file);

                Console.WriteLine(fileType);

                if (fileType == ".png" || fileType == ".jpg" || fileType == ".gif" || fileType == ".docx" || fileType == ".css")
                    continue;

                var fileDir = System.IO.Path.GetDirectoryName(file);
                Console.WriteLine(fileDir);

                var docFile = System.IO.Path.Combine(fileDir, fileName + ".docx");

                System.IO.File.Copy(TemPath, docFile, true);


                Report report = new Report();
                report.CreateNewDocument(docFile); //模板路径

                string text = File.ReadAllText(file);

                Console.Write(text.Length);

                report.InsertText("code", text);

                report.Document.Close();

                //try
                //{
                //    report.SaveDocument(docFile); //文档路径
                //}
                //catch (Exception)
                //{

                //} 
            }


            Console.WriteLine("OK");

            //report.SaveDocument(RepPath); //文档路径

        }
    }
}

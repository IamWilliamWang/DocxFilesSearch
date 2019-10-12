using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocxFilesSearch
{
    class SearchWordsText
    {
        /// <summary>
        /// 循环Trim
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        private String Trim(String str)
        {
            while (true)
            {
                string newStr = str.Trim();
                if (newStr == str)
                    break;
                str = newStr;
            }
            return str;
        }

        private String lastFileName = ""; // 储存上一次显示时的文件名
        /// <summary>
        /// 使用StringBuilder保存输出详情
        /// </summary>
        /// <param name="printer">用于保存的StringBuilder</param>
        /// <param name="fullFileName">文件全名</param>
        /// <param name="lines">全文内容</param>
        /// <param name="foundRowIndex">需要输出的行数信息</param>
        private void PrintDetails(ref StringBuilder printer, ref String fullFileName, ref String[] lines, int foundRowIndex)
        {
            if (fullFileName != lastFileName) 
            {
                if (printer.Length != 0)
                {
                    printer.AppendLine();
                    printer.AppendLine();
                }
                printer.AppendLine("Matched: " + fullFileName);
                printer.AppendLine("------------------------------------------------");
                lastFileName = fullFileName;
            }
            printer.AppendFormat("Line {0,3}: {1}\r\n", foundRowIndex + 1, lines[foundRowIndex]);
        }

        /// <summary>
        /// 不使用StringBuilder直接输出详情到屏幕中
        /// </summary>
        /// <param name="fullFileName">文件全名</param>
        /// <param name="lines">全文内容</param>
        /// <param name="foundRowIndex">需要输出的行数信息</param>
        private void PrintDetailsNoCache(ref String fullFileName, ref String[] lines, int foundRowIndex)
        {
            if (fullFileName != lastFileName)
            {
                if (lastFileName != "")
                {
                    Console.WriteLine();
                    Console.WriteLine();
                }
                Console.WriteLine("Matched: " + fullFileName);
                Console.WriteLine("------------------------------------------------");
                lastFileName = fullFileName;
            }
            Console.WriteLine("Line {0,3}: {1}\r\n", foundRowIndex + 1, lines[foundRowIndex]);
        }

        /// <summary>
        /// 刷新缓存，将StringBuilder所有内容输出到屏幕中
        /// </summary>
        /// <param name="buffer">存储输出文字的StringBuilder</param>
        private void FlushBuffer(ref StringBuilder buffer)
        {
            string result = buffer.ToString();
            if (result != "")
                Console.WriteLine(result);
            else
                Console.WriteLine("没有匹配结果！");
        }

        private void CheckEmpty()
        {
            if (lastFileName == "")
                Console.WriteLine("没有匹配结果！");
        }

        public void Run(bool useBuffer = true)
        {
            // 输入必要的信息
            Console.Write("请输入要查询Word文档的文件夹：");
            String folder = Console.ReadLine().Replace("\"", "").Trim();
            Console.Write("请输入要查询的内容（不要包含换行）：");
            String searchingContent = Console.ReadLine().Trim();

            StringBuilder printBuilder = new StringBuilder();
            // 开始搜索
            var directory = new DirectoryInfo(folder);
            foreach (var fileInfo in directory.GetFiles())
            {
                if (fileInfo.Extension != ".docx")
                    continue;
                String fullFileName = fileInfo.FullName;
                Word thisWordDocument = new Word(fullFileName);
                if (thisWordDocument.ReadWord().Contains(searchingContent))
                {
                    String[] docxLines = thisWordDocument.ReadWordLines();
                    for (var rowIndex = 0; rowIndex < docxLines.Length; rowIndex++)
                    {
                        if (docxLines[rowIndex].Contains(searchingContent))
                        {
                            if (useBuffer)
                                PrintDetails(ref printBuilder, ref fullFileName, ref docxLines, rowIndex);
                            else
                                PrintDetailsNoCache(ref fullFileName, ref docxLines, rowIndex);
                        }
                    }
                }
            }
            if (useBuffer)
                FlushBuffer(ref printBuilder);
            else
                CheckEmpty();
        }

        static void Main(string[] args)
        {
            new SearchWordsText().Run(false);
            Console.ReadKey();
        }
    }
}

﻿using NPOI.HPSF;
using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocxFilesSearch
{
    class Word
    {
        private string filename;
        private XWPFDocument docWrite = new XWPFDocument();
        public string Font { get; set; } = "黑体";
        public int FontSize { get; set; } = 13;
        public int Length { get { return this.ReadWord().Replace("\r", "").Replace("\n", "").Length; } }
        public Word(string docxFileName = "document.docx", string authorName = "")
        {
            docWrite.GetProperties().CoreProperties.Creator = authorName;
            this.filename = docxFileName;
        }

        public void WriteDocx(string content)
        {
            string[] contentLines = content.Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries); //提取每一行的内容
            WriteDocx(contentLines);
        }

        public void WriteDocx(string[] contentLines)
        {
            foreach (string line in contentLines)
            {
                XWPFParagraph paragraph = docWrite.CreateParagraph(); //每一行对应word中的一段
                paragraph.Alignment = ParagraphAlignment.BOTH;
                XWPFRun run = paragraph.CreateRun(); //给这一段插入一行文字内容
                run.SetText(line);
                run.FontFamily = Font;
                run.FontSize = FontSize;
            }
            try
            {
                using (FileStream outStream = new FileStream(filename, FileMode.Create))
                    docWrite.Write(outStream);
            }
            catch (DirectoryNotFoundException)
            {
                throw;
            }
        }

        public String ReadWord()
        {
            StringBuilder textReaded = new StringBuilder();
            String[] contentLines = ReadWordLines();
            foreach (String line in contentLines)
                textReaded.AppendLine(line);
            return textReaded.ToString();
        }

        public String[] ReadWordLines()
        {
            List<String> lines = new List<string>();
            using (FileStream docxStream = File.OpenRead(this.filename))
            {
                XWPFDocument docRead = new XWPFDocument(docxStream);
                if (docRead.Paragraphs.Count == 0)
                    return lines.ToArray();
                foreach (XWPFParagraph paragraph in docRead.Paragraphs)
                {
                    string paragraphText = paragraph.ParagraphText; //获得该段的文本，因为不需要管文字格式所以不用获取XWPFRun
                    lines.Add(paragraphText);
                }
                this.FontSize = docRead.Paragraphs[0].Runs[0].FontSize;
                this.Font = docRead.Paragraphs[0].Runs[0].FontFamily;
            }
            return lines.ToArray();
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Drawing;
using System.Threading;
namespace MSWord_VWS_Documentation_Export
{

    class Program
    {
        static string OutputDir = "";//"c:\\tutorial export\\";
        static string InputFile = "";//"c:\\users\\chadwickr\\downloads\\Building a Game in the Virtual World Sandbox.docx";
        static string OutputFile = "export.html";
        static bool processBolds = true;
        static int ImageCount = 0;
        class ContentObject
        {
            public int start;
            object content;
            int type;
            Microsoft.Office.Interop.Word.Application WordApp;
            public ContentObject(int s, object c, int t, Microsoft.Office.Interop.Word.Application a)
            {
                this.start = s;
                this.content = c;
                this.type = t;
                this.WordApp = a;
            }
            private void SaveToImage(string filePath)
            {
                ((InlineShape)this.content).Select();
                this.WordApp.Selection.CopyAsPicture();


                IDataObject data = null;
                Exception threadEx = null;
                ThreadStart staThreads = new ThreadStart(
                    delegate
                    {
                        try
                        {
                            data = Clipboard.GetDataObject();
                            if (data != null && data.GetDataPresent(typeof(Bitmap)))
                            {
                                Bitmap image = (Bitmap)data.GetData(typeof(Bitmap));
                                image.Save(filePath);
                            }
                        }

                        catch (Exception ex)
                        {
                            threadEx = ex;
                        }
                    });
                Thread staThread = new Thread(staThreads);
                staThread.SetApartmentState(ApartmentState.STA);
                staThread.Start();
                staThread.Join();




            }
            public static string cleanString(string text)
            {
                var s = text;
                // smart single quotes and apostrophe
                s = Regex.Replace(s, "[\u2018|\u2019|\u201A]", "'");
                // smart double quotes
                s = Regex.Replace(s, "[\u201C|\u201D|\u201E]", "\"");
                // ellipsis
                s = Regex.Replace(s, "\u2026", "...");
                // dashes
                s = Regex.Replace(s, "[\u2013|\u2014]", "-");
                // circumflex
                s = Regex.Replace(s, "\u02C6", "^");
                // open angle bracket
                s = Regex.Replace(s, "\u2039", "<");
                // close angle bracket
                s = Regex.Replace(s, "\u203A", ">");
                // spaces
                s = Regex.Replace(s, "[\u02DC|\u00A0]", " ");
                s = Regex.Replace(s, "[\u0001|\u0015]", " ");
                s = Regex.Replace(s, @"[^\u0000-\u007F]", string.Empty);
                s = s.Trim();



                if (s == "/")
                    s = "";
                return s;
            }
            public static string dobolds(Range r)
            {
                if(!processBolds)
                    return r.Text;

                List<Tuple<int, int>> ranges = new List<Tuple<int, int>>();
                bool inbold = false;
                int start = -1;
                foreach (Microsoft.Office.Interop.Word.Range rngWord in r.Words)
                {

                    if (rngWord.Bold != 0)
                    {
                        if (rngWord.Text.Trim().Length > 0 && !inbold)
                        {
                            inbold = true;
                            start = rngWord.Start;

                        }
                    }
                    else
                    {
                        if (inbold)
                        {
                            ranges.Add(new Tuple<int, int>(start, rngWord.Start));
                            inbold = false;
                        }
                    }

                }

                var text = r.Text;

                foreach (Tuple<int, int> i in ranges)
                {

                    r.SetRange(i.Item1, i.Item2);
                    var bold = r.Text;
                    Console.WriteLine(bold);

                    text = text.Replace(bold, "<span class='bold'>" + bold + "</span>");

                }
                return text;
            }
            public void print(System.IO.StreamWriter outfile)
            {
                if (this.type == 0)
                {
                    Paragraph p = ((Paragraph)(this.content));
                    String text = p.Range.Text;
                    text = cleanString(dobolds(((Paragraph)(this.content)).Range));
                    if (text.Length == 0) return;
                    var fontsize = "null";
                    if (p.Range.Words != null && p.Range.Words.Count > 0)
                        fontsize = Math.Floor(p.Range.Words[1].Font.Size).ToString();
                    var indent = (p.LeftIndent);

                    var liststring = p.Range.ListFormat.ListString;
                    if (liststring.Length == 1)
                        liststring = "";
                    else
                        liststring = liststring + " ";
                    string listclass = "numberedlist";
                    if (liststring == " ")
                        listclass = "bullet";

                    if (listclass == "numberedlist")
                    {
                        outfile.WriteLine("<div><div class='" + listclass + "'>" + liststring + "</div>");
                    }
                    outfile.WriteLine("<div class='step font" + fontsize + " indent" + indent + "'>" + text + "</div>");
                    if (listclass == "numberedlist")
                    {
                        outfile.WriteLine("</div>");
                    }

                }
                if (this.type == 1)
                {
                    // outfile.WriteLine(((Paragraph)(this.content)).Range.Start);

                    string filename = "Image" + ImageCount.ToString() + ".png";// System.IO.Path.GetRandomFileName();
                    ImageCount++;
                    
                    this.SaveToImage(System.IO.Path.Combine(System.IO.Path.Combine(OutputDir, "images\\") + filename));

                    outfile.WriteLine("<img class='image' src='" + "images\\" +filename + "' />\n");

                }
                if (this.type == 2)
                {
                    Shape s = ((Shape)(this.content));
                    String text = cleanString(dobolds(s.TextFrame.TextRange));

                    string fill = "Color" + s.Fill.ForeColor.RGB.ToString();
                    if (text.Split('\r').Length > 1)
                        outfile.Write("<div class='code'>");
                    else
                        outfile.WriteLine("<div class='inset " + fill + "'>");
                    outfile.WriteLine(text);
                    outfile.WriteLine("</div>");

                }
            }
        }

        static void WriteHead(System.IO.StreamWriter outfile)
        {

            string head = "<!DOCTYPE html>\n" +

            "<html>\n" +
              "<head>\n" +
               "<LINK href='style.css' rel='stylesheet' type='text/css'>"+
                "</head>\n" +
              "</head>\n" +
              "<body>\n" +
                  "<div class='article'>\n";

            outfile.WriteLine(head);

        }
        static void WriteFoot(System.IO.StreamWriter outfile)
        {

            string head = 
                "</div>\n" + 
              "</body>\n"+
            "</html>\n";
            outfile.WriteLine(head);

        }
        static void Main(string[] args)
        {
     
            for (int i = 0; i < args.Length; i++)
            {
                if (args[i] == "-o")
                    OutputFile = args[i + 1];
                if (args[i] == "-i")
                    InputFile = args[i + 1];
                 if (args[i] == "-d")
                    OutputDir = args[i + 1];
                 if (args[i] == "-b")
                     processBolds = Boolean.Parse(args[i + 1]);

                
            }

            object oMissing = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Word.Application WordApp = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document WordDoc = WordApp.Documents.Open(InputFile,ref oMissing, true);
            List<ContentObject> contents = new List<ContentObject>();

            for (var i = 1; i < WordDoc.Paragraphs.Count; i++)
            {
                   Paragraph p = WordDoc.Paragraphs[i];
                
                 //  Console.WriteLine(p.Range.Text);
                 //  Console.WriteLine(p.Range.Start);
                   contents.Add(new ContentObject(p.Range.Start, p, 0, WordApp));
            }
            for (var i = 1; i < WordDoc.InlineShapes.Count; i++)
            {
                InlineShape s = WordDoc.InlineShapes[i];
                contents.Add(new ContentObject(s.Range.Start, s, 1, WordApp));
              //  Console.WriteLine(s.Range.Text);
              //  Console.WriteLine(s.Range.Start);
            }
            for (var i = 1; i < WordDoc.Shapes.Count; i++)
            {
               Shape s = WordDoc.Shapes[i]; 
             //  Console.WriteLine(s.TextFrame.TextRange.Text);
             //  Console.WriteLine(s.Anchor.Start);
               contents.Add(new ContentObject(s.Anchor.Start, s, 2, WordApp));
               
            }


            contents.Sort(delegate(ContentObject a, ContentObject b)
            {
                if (a == null) return -1;
                if (a.start < b.start) return -1;
                else return 1;
            });

            if(!System.IO.Directory.Exists(OutputDir))
            {
                System.IO.Directory.CreateDirectory(OutputDir);
            }
            if (!System.IO.Directory.Exists(System.IO.Path.Combine(OutputDir, "images\\")))
            {
                System.IO.Directory.CreateDirectory(System.IO.Path.Combine(OutputDir, "images\\"));
            }
            System.IO.StreamWriter file = new System.IO.StreamWriter(System.IO.Path.Combine( OutputDir , OutputFile));

            WriteHead(file);
            for (int i = 0; i < contents.Count; i++)
            {
                contents[i].print(file);

            }
            WriteFoot(file);
            object doNotSaveChanges = Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges;
            WordDoc.Close(ref doNotSaveChanges);
           
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.IO;
using System.Diagnostics;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;

namespace AttributeConverter
{
    class WordProcessor : IProcessor
    {

        private readonly Config config;

        public WordProcessor(Config config)
        {
            this.config = config;
        }

        public void WriteDocument(Table[] content)
        {
            DateTime startTime = DateTime.Now;
            Stopwatch sw = new Stopwatch();
            sw.Start();
            long currentRuntime = 0;
            Stopwatch swSingle = new Stopwatch();

            int length = content.Length;

            FileInfo outputFile = new FileInfo(config.outputfile);
            File.Delete(outputFile.FullName);

            object pageBreak = Word.WdBreakType.wdPageBreak;

            Console.WriteLine();
            Console.WriteLine("Starte Word Verarbeitung...");

            object oMissing = System.Reflection.Missing.Value;
            object oEndOfDoc = "\\endofdoc"; /* \endofdoc is a predefined bookmark */

            //Start Word and create a new document.
            Word._Application oWord = null;
            Word._Document oDoc = null;
            try
            {
                oWord = new Word.Application
                {
                    Visible = false
                };

                Word.Range myRange = null;
                if (config.wordTemplate != null)
                {
                    Console.WriteLine("Lade Word Template: " + config.wordTemplate);

                    oDoc = oWord.Documents.Open(config.wordTemplate);

                    // find the content string to replace it with the actual tables
                    int pgs = oDoc.Paragraphs.Count;
                    for (int j=0; j<pgs; j++)
                    {
                        string pgText = oDoc.Paragraphs[j + 1].Range.Text;
                        if (pgText.StartsWith("CONTENT"))
                        {
                            myRange = oDoc.Paragraphs[j + 1].Range;
                            myRange.Select();
                            break;
                        }
                    }
                } else
                {
                    // if no template is defined
                    oDoc = oWord.Documents.Add(ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing);
                    myRange = oDoc.Range(oDoc.Content.End - 1, ref oMissing);
                }

                Console.WriteLine("Schreibe Datei: " + outputFile.FullName);
                Console.WriteLine();

                if (myRange == null)
                {
                    // fail-safe if no range has been found yet
                    myRange = oDoc.Range(oDoc.Content.End - 1, ref oMissing);
                }

                string previousCategory = "";

                // table creation
                double runTime = 0.0;
                int i = 0;
                for (; i < length; i++)
                {
                    swSingle.Restart();
                    try
                    {
                        myRange.InsertParagraph();
                        myRange = oDoc.Range(myRange.End - 1, myRange.End);

                        // add a paragraph after the table and select it for replacement with follow-up content
                        for (int r = 0; r < content[i].entries.Count; r++)
                        {
                            if (content[i].entries[r].heading != Word.WdBuiltinStyle.wdStyleNormal)
                            {
                                // headline formatting
                                if (Word.WdBuiltinStyle.wdStyleHeading1 == content[i].entries[r].heading && !previousCategory.Equals(content[i].entries[r].value.Replace('\n', ' ')))
                                {
                                    previousCategory = content[i].entries[r].value.Replace('\n', ' ');
                                    myRange.InsertParagraph();
                                    myRange.Text = previousCategory;
                                    myRange.set_Style(content[i].entries[r].heading);
                                    myRange.InsertParagraphAfter();
                                    myRange.InsertParagraphAfter();
                                    myRange = oDoc.Range(myRange.End - 1, myRange.End);
                                    myRange.Select();
                                }
                                else if (Word.WdBuiltinStyle.wdStyleHeading1 != content[i].entries[r].heading)
                                {
                                    myRange.InsertParagraph();
                                    myRange.Text = content[i].entries[r].value.Replace('\n', ' ');
                                    myRange.set_Style(content[i].entries[r].heading);
                                    myRange.InsertParagraphAfter();
                                    myRange.InsertParagraphAfter();
                                    myRange = oDoc.Range(myRange.End - 1, myRange.End);
                                    myRange.Select();
                                }
                            }
                        }
                        myRange.set_Style(Word.WdBuiltinStyle.wdStyleNormal);
                        myRange.InsertParagraphAfter();
                        myRange.InsertParagraphAfter();
                        myRange.Select();

                        // add a pre-defined table with columns and rows
                        Word.Table oTable = oDoc.Tables.Add(myRange, content[i].entries.Count, 2, ref oMissing, ref oMissing);
                        oTable.Range.Borders.Enable = 1;
                        oTable.Range.ParagraphFormat.SpaceAfter = 6;
                        oTable.Range.set_Style(Word.WdBuiltinStyle.wdStyleNormal);

                        for (int r = 0; r < content[i].entries.Count; r++)
                        {
                            // the actual cell text
                            oTable.Cell(r + 1, 1).Range.Text = content[i].entries[r].name.Replace('\n', ' ');
                            oTable.Cell(r + 1, 2).Range.Text = content[i].entries[r].value;

                            if (content[i].entries[r].header)
                            {
                                // table cell styling
                                oTable.Cell(r + 1, 1).Range.Font.Bold = 1;
                                oTable.Cell(r + 1, 1).Range.Shading.BackgroundPatternColor = WdColor.wdColorBlueGray;
                                oTable.Cell(r + 1, 2).Range.Shading.BackgroundPatternColor = WdColor.wdColorBlueGray;
                            }
                            
                            if (content[i].entries[r].heading != Word.WdBuiltinStyle.wdStyleNormal)
                            {
                                oTable.Cell(r + 1, 2).Range.Text = content[i].entries[r].value.Replace('\n', ' ');
                                /*
                                // headline formatting
                                if (Word.WdBuiltinStyle.wdStyleHeading1 == content[i].entries[r].heading && !previousCategory.Equals(content[i].entries[r].value.Replace('\n', ' ')))
                                {
                                    previousCategory = content[i].entries[r].value.Replace('\n', ' ');
                                    oTable.Cell(r + 1, 2).Range.Text = previousCategory;
                                    oTable.Cell(r + 1, 2).Range.set_Style(content[i].entries[r].heading);
                                } 
                                else if (Word.WdBuiltinStyle.wdStyleHeading1 != content[i].entries[r].heading)
                                {
                                    oTable.Cell(r + 1, 2).Range.Text = content[i].entries[r].value.Replace('\n', ' ');
                                    oTable.Cell(r + 1, 2).Range.set_Style(content[i].entries[r].heading);
                                }
                            */
                            }
                        }

                        // add a paragraph after the table and select it for replacement with follow-up content
                        myRange = oDoc.Range(oTable.Range.End, oTable.Range.End);
                        myRange.set_Style(Word.WdBuiltinStyle.wdStyleNormal);
                        myRange.InsertParagraphAfter();
                        myRange.Select();

                        if (i + 1 < length)
                        {
                            // add a page-break between tables
                            Word.Paragraph oPara4;
                            oPara4 = oDoc.Content.Paragraphs.Add(myRange);
                            oPara4.Range.InsertBreak(WdBreakType.wdPageBreak);
                            myRange = oDoc.Range(oPara4.Range.End - 1, oPara4.Range.End);
                            myRange.set_Style(Word.WdBuiltinStyle.wdStyleNormal);
                            myRange.Select();
                        }

                        // runtime calculation
                        swSingle.Stop();
                        currentRuntime += swSingle.ElapsedMilliseconds;
                        runTime = (currentRuntime / 1000) / (i + 1) * length;
                        Console.Write("\r{0}%  {1}/{2}   Laufzeit: {3}                 ", i * 100 / length, i, length, new DateTime(startTime.Ticks).AddSeconds(runTime).ToShortTimeString());
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("FATAL: " + ex.Message);
                    }
                }
            }
            finally
            {
                // safe-close the file and word
                if (oDoc != null)
                {
                    oDoc.SaveAs2(outputFile.FullName);
                    oDoc.Close();
                }
                if (oWord != null)
                {
                    oWord.Quit(oMissing, oMissing, oMissing);
                }
            }

            sw.Stop();
            long processingTime = sw.ElapsedMilliseconds;
            Console.Write("\r{0}%                                                        ", 100);
            Console.WriteLine();
            Console.WriteLine("Word Dokument erzeugt in " + (processingTime / 1000) + "s mit " + length + " Seiten.");
        }

    }
}

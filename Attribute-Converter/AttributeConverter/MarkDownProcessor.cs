using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.IO;
using System.Diagnostics;
using Word = Microsoft.Office.Interop.Word;

namespace AttributeConverter
{
    class MarkDownProcessor : IProcessor
    {

        private readonly Config config;

        public MarkDownProcessor(Config config)
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

            object pageBreak = "\r\n";

            Console.WriteLine();
            Console.WriteLine("Starte Markdown Verarbeitung...");

            FileStream fileStream = outputFile.Create();
            StreamWriter streamWriter = new StreamWriter(fileStream);

            //Start Word and create a new document.
            try
            {

                Console.WriteLine("Schreibe Datei: " + outputFile.FullName);
                Console.WriteLine();

                streamWriter.WriteLine("");
                streamWriter.WriteLine("<link href=\"custom.md.css\" rel=\"stylesheet\"></link>");
                streamWriter.WriteLine("");
                streamWriter.WriteLine("");

                string toc = Languages.DE.Equals(config.language) ? "Inhaltsverzeichnis" : Util.GetConfigTranslation(config, "Inhaltsverzeichnis", config.language);
                streamWriter.WriteLine("# " + toc);

                string previousCategory = "";

                // table creation
                double runTime = 0.0;

                int i = 0;
                for (int j = 1, k = 0; i < length; i++)
                {
                    // add a paragraph after the table and select it for replacement with follow-up content
                    for (int r = 0; r < content[i].entries.Count; r++)
                    {
                        if (content[i].entries[r].heading != Word.WdBuiltinStyle.wdStyleNormal)
                        {
                            // headline formatting
                            if (Word.WdBuiltinStyle.wdStyleHeading1 == content[i].entries[r].heading && !previousCategory.Equals(content[i].entries[r].value.Replace('\n', ' ')))
                            {
                                k++;
                                previousCategory = content[i].entries[r].value.Replace('\n', ' ');
                                streamWriter.WriteLine(k.ToString() + ". [" + previousCategory + "](#" + ConvertToLink(k.ToString() + " " + previousCategory) + ")");
                                j = 1;
                            }
                            else if (Word.WdBuiltinStyle.wdStyleHeading1 != content[i].entries[r].heading)
                            {
                                streamWriter.WriteLine("    " + j.ToString() + ". [" + content[i].entries[r].value.Replace('\n', ' ') + "](#" + ConvertToLink(j.ToString() + " " + content[i].entries[r].value.Replace('\n', ' ')) + ")");
                                j++;
                            }
                        }
                    }
                }
                
                streamWriter.WriteLine(pageBreak);

                i = 0;
                for (int j = 1, k = 0; i < length; i++)
                {
                    swSingle.Restart();
                    try
                    {

                        // add a paragraph after the table and select it for replacement with follow-up content
                        for (int r = 0; r < content[i].entries.Count; r++)
                        {
                            if (content[i].entries[r].heading != Word.WdBuiltinStyle.wdStyleNormal)
                            {
                                // headline formatting
                                if (Word.WdBuiltinStyle.wdStyleHeading1 == content[i].entries[r].heading && !previousCategory.Equals(content[i].entries[r].value.Replace('\n', ' ')))
                                {
                                    k++;
                                    previousCategory = content[i].entries[r].value.Replace('\n', ' ');
                                    streamWriter.WriteLine("# " + k.ToString() + ". " + previousCategory);
                                    j = 1;
                                }
                                else if (Word.WdBuiltinStyle.wdStyleHeading1 != content[i].entries[r].heading)
                                {
                                    streamWriter.WriteLine("## " + j.ToString() + ". "  + content[i].entries[r].value.Replace('\n', ' '));
                                    j++;
                                }
                            }
                        }

                        // add a pre-defined table with columns and rows
                        bool headerWritten = false;

                        for (int r = 0; r < content[i].entries.Count; r++)
                        {
                            string formatter = "";
                            string formatter2 = "";

                            if (content[i].entries[r].header)
                            {
                                // table cell styling
                                formatter = "**";
                                formatter2 = "**";
                            }

                            string value = content[i].entries[r].value;

                            if (content[i].entries[r].heading != Word.WdBuiltinStyle.wdStyleNormal)
                            {
                                value = content[i].entries[r].value.Replace('\n', ' ');
                            }

                            if (value == null || value.Length == 0)
                            {
                                value = "";
                                formatter2 = "";
                            }

                            // the actual cell text
                            streamWriter.WriteLine("|" + formatter + content[i].entries[r].name.Replace('\n', ' ') + formatter + "|" + formatter2 + value + formatter2 + "|");

                            if (!headerWritten && content[i].entries[r].header)
                            {
                                streamWriter.WriteLine("|:---:|---|");
                                headerWritten = true;
                            }

                        }

                        // add a paragraph after the table and select it for replacement with follow-up content
                        string top = Languages.DE.Equals(config.language) ? "Anfang" : Util.GetConfigTranslation(config, "Anfang", config.language);
                        streamWriter.WriteLine("["+ top + "](#" + ConvertToLink(toc) + ")");
                        streamWriter.WriteLine(pageBreak);

                        streamWriter.Flush();

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

                string created = Languages.DE.Equals(config.language) ? "Generiert" : Util.GetConfigTranslation(config, "Generiert", config.language);
                streamWriter.WriteLine("*" + created + ": " + DateTime.Now.ToShortDateString() + "*");
            }
            finally
            {
                if (streamWriter != null)
                {
                    streamWriter.Flush();
                }
                if (fileStream != null)
                {
                    fileStream.Flush();
                    fileStream.Dispose();
                }
            }

            sw.Stop();
            long processingTime = sw.ElapsedMilliseconds;
            Console.Write("\r{0}%                                                        ", 100);
            Console.WriteLine();
            Console.WriteLine("MarkDown Dokument erzeugt in " + (processingTime / 1000) + "s mit " + length + " Seiten.");
        }

        private string ConvertToLink(string input)
        {
            return input.ToLower().Replace(" ", "-").Replace("(", "").Replace(")", "");
        }
    }
}

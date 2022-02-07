using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.IO;
using System.Diagnostics;
using ExcelDataReader;
using Word = Microsoft.Office.Interop.Word;

namespace AttributeConverter
{

    class ExcelParser
    {
        private readonly Config config;

        public ExcelParser(Config config)
        {
            this.config = config;
        }


        public Table[] ProcessExcelFile()
        {
            Stopwatch sw = new Stopwatch();
            sw.Start();

            List<Table> contents = new List<Table>();
            Console.WriteLine("Lade Exceldatei: " + config.excelFile);
            FileInfo inputFile = new FileInfo(config.excelFile);

            using (var stream = File.Open(inputFile.FullName, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet();

                    foreach (System.Data.DataTable tableResult in result.Tables)
                    {
                        // switch the excel tab to Attributes
                        if (tableResult.TableName == config.workSheetAttributes)
                        {
                            // initial data
                            List<int> columnsToParse = new List<int>();
                            int colBlockAttributeRanges = config.columnStartBlockRules - config.columnStartBlockAttribute;
                            int colBlockRulesRange = 100;
                            int indexRow = -1;
                            object[] headers = null;
                            object[] subHeaders = null;
                            int headerRow = -1;

                            // per-row changeable data
                            string lastCategory = null;
                            string displayCategory = null;
                            string lastSubCategory = null;
                            string displaySubCategory = null;
                            string lastBlockRule = "";

                            for (int i = 0; i < tableResult.Rows.Count; i++)
                            {
                                object[] data = tableResult.Rows[i].ItemArray;
                                int textLength = data[0] is string ? ((string)data[0]).Length : 0;

                                if (i == config.outputMarkerRow)
                                {
                                    for (int j = 0; j < data.Length; j++)
                                    {
                                        if ("x".Equals(data[j].ToString()))
                                        {
                                            columnsToParse.Add(j);
                                        }
                                    }
                                }

                                // scan through the index contents and get the next category and block rule
                                if (headers != null && headers.Length > 0 && config.columnCategory > -1 && config.columnSubCategory > -1)
                                {
                                    if (data[config.columnCategory] is string && ((string)data[config.columnCategory]).Length > 0)
                                    {
                                        lastCategory = (string)data[config.columnCategory];
                                        displayCategory = GetTranslation(result.Tables, lastCategory, "", "", config.language, false);
                                    }
                                    if (data[config.columnSubCategory] is string && ((string)data[config.columnSubCategory]).Length > 0)
                                    {
                                        lastSubCategory = (string)data[config.columnSubCategory];
                                        displaySubCategory = GetTranslation(result.Tables, lastCategory, lastSubCategory, "", config.language, false);

                                        // check the block rule as this is the same length as sub category
                                        if (config.columnStartBlockRules > -1 && data[config.columnStartBlockRules + 1] is string && ((string)data[config.columnStartBlockRules + 1]).Length > 0)
                                        {
                                            lastBlockRule = (string)data[config.columnStartBlockRules + 1];
                                        }
                                    }
                                }

                                if (textLength > 0)
                                {
                                    // the first header row
                                    if (indexRow == -1)
                                    {
                                        indexRow = i;
                                    }
                                    if (textLength > 1 && headers == null)
                                    {
                                        headers = data;
                                        subHeaders = tableResult.Rows[i - 1].ItemArray;
                                        headerRow = i;
                                    }
                                    else if (config.columnAttribute > -1 && data[config.columnAttribute].ToString().Length > 0 && (
                                        data[0].ToString().ToLower().Equals("x")
                                        || data[0].ToString().ToLower().Equals("") && data[1].ToString().ToLower().Equals("")
                                        ))
                                    {
                                        // Fill data for extraction

                                        Table table = new Table();
                                        table.entries = new List<TableEntry>();
                                        contents.Add(table);

                                        // Structure level attribute
                                        table.entries.Add(new TableEntry()
                                        {
                                            name = GetHeaderTranslation(result.Tables, headers[config.columnStructureLevel].ToString(), headerRow, config.columnStructureLevel, config.language),
                                            value = "",
                                            header = true
                                        });
                                        // Category attribute
                                        table.entries.Add(new TableEntry()
                                        {
                                            name = GetHeaderTranslation(result.Tables, headers[config.columnCategory].ToString(), headerRow, config.columnCategory, config.language),
                                            value = displayCategory,
                                            heading = Word.WdBuiltinStyle.wdStyleHeading1
                                        });
                                        // Subcategory attribute
                                        table.entries.Add(new TableEntry()
                                        {
                                            name = GetHeaderTranslation(result.Tables, headers[config.columnSubCategory].ToString(), headerRow, config.columnSubCategory, config.language),
                                            value = displaySubCategory
                                        });
                                        // Fieldname attribute
                                        table.entries.Add(new TableEntry()
                                        {
                                            name = GetHeaderTranslation(result.Tables, headers[config.columnAttribute].ToString(), headerRow, config.columnAttribute, config.language),
                                            value = GetTranslation(result.Tables, lastCategory, lastSubCategory, data[config.columnAttribute].ToString(), config.language, false),
                                            heading = Word.WdBuiltinStyle.wdStyleHeading2
                                        });
                                        // Custom description attribute
                                        table.entries.Add(new TableEntry()
                                        {
                                            name = Languages.DE.Equals(config.language) ? config.columnDescription : Util.GetConfigTranslation(config, config.columnDescription, config.language),
                                            value = GetTranslation(result.Tables, lastCategory, lastSubCategory, data[config.columnAttribute].ToString(), config.language, true),
                                        });
                                        // Custom path attribute
                                        table.entries.Add(new TableEntry()
                                        {
                                            name = Languages.DE.Equals(config.language) ? config.columnPath : Util.GetConfigTranslation(config, config.columnPath, config.language),
                                            value = GetKey(result.Tables, lastCategory, lastSubCategory, data[config.columnAttribute].ToString(), config.language, false),
                                        });

                                        // Components
                                        table.entries.Add(new TableEntry()
                                        {
                                            name = GetHeaderTranslation(result.Tables, headers[config.columnStartBlockComponent].ToString(), headerRow, config.columnStartBlockComponent, config.language),
                                            value = "",
                                            header = true
                                        });
                                        for (int d = config.columnStartBlockComponent; d < config.columnStartBlockComponent + config.columnBlockComponentRange; d++)
                                        {
                                            if (columnsToParse.Contains(d) && headers[d].ToString().Length > 0 && (!config.skipEmtpyLines || config.skipEmtpyLines && data[d].ToString().Length > 0))
                                            {
                                                table.entries.Add(new TableEntry()
                                                {
                                                    name = GetHeaderTranslation(result.Tables, headers[d].ToString(), headerRow, d, config.language),
                                                    value = data[d].ToString()
                                                });
                                            }
                                        }
                                        table.entries.Add(new TableEntry()
                                        {
                                            name = GetHeaderTranslation(result.Tables, headers[config.columnStartBlockAttribute].ToString(), headerRow, config.columnStartBlockAttribute, config.language),
                                            value = "",
                                            header = true
                                        });
                                        for (int d = config.columnStartBlockAttribute; d < config.columnStartBlockAttribute + colBlockAttributeRanges; d++)
                                        {
                                            if (columnsToParse.Contains(d) && headers[d].ToString().Length > 0 && (!config.skipEmtpyLines || config.skipEmtpyLines && data[d].ToString().Length > 0))
                                            {
                                                string value = data[d].ToString();

                                                if (headers[d].ToString().StartsWith(config.columnValueLists))
                                                {
                                                    value = GetValueList(result.Tables, value, config.language);
                                                }

                                                table.entries.Add(new TableEntry()
                                                {
                                                    name = GetHeaderTranslation(result.Tables, headers[d].ToString(), headerRow, d, config.language),
                                                    value = value
                                                });
                                            }
                                        }
                                        table.entries.Add(new TableEntry()
                                        {
                                            name = GetHeaderTranslation(result.Tables, headers[config.columnStartBlockRules].ToString(), headerRow, config.columnStartBlockRules, config.language),
                                            value = "",
                                            header = true
                                        });
                                        if (lastBlockRule.Length > 0)
                                        {
                                            table.entries.Add(new TableEntry()
                                            {
                                                name = GetHeaderTranslation(result.Tables, headers[config.columnStartBlockRules].ToString(), headerRow, config.columnStartBlockRules, config.language),
                                                value = lastBlockRule
                                            });
                                        }
                                        for (int d = config.columnStartBlockRules + 1; d < config.columnStartBlockRules + colBlockRulesRange; d++)
                                        {
                                            if (columnsToParse.Contains(d) && headers[d].ToString().Length > 0 && (!config.skipEmtpyLines || config.skipEmtpyLines && data[d].ToString().Length > 0))
                                            {
                                                string header = GetHeaderTranslation(result.Tables, headers[d].ToString(), headerRow, d, config.language);
                                                if (d >= config.columnStartBlockRules + 2 && d <= config.columnStartBlockRules + 4)
                                                {
                                                    header = GetHeaderTranslation(result.Tables, subHeaders[d].ToString(), headerRow - 1, d, config.language) + " " + header;
                                                }
                                                if (d >= config.columnStartBlockRules + 5 && d <= config.columnStartBlockRules + 7)
                                                {
                                                    header = GetHeaderTranslation(result.Tables, subHeaders[d].ToString(), headerRow - 1, d, config.language) + " " + header;
                                                }
                                                if (d >= config.columnStartBlockRules + 8 && d <= config.columnStartBlockRules + 10)
                                                {
                                                    header = GetHeaderTranslation(result.Tables, subHeaders[d].ToString(), headerRow - 1, d, config.language) + " " + header;
                                                }
                                                if (d >= config.columnStartBlockRules + 11 && d <= config.columnStartBlockRules + 13)
                                                {
                                                    header = GetHeaderTranslation(result.Tables, subHeaders[d].ToString(), headerRow - 1, d, config.language) + " " + header;
                                                }
                                                if (d >= config.columnStartBlockRules + 14 && d <= config.columnStartBlockRules + 16)
                                                {
                                                    header = GetHeaderTranslation(result.Tables, subHeaders[d].ToString(), headerRow - 1, d, config.language) + " " + header;
                                                }
                                                table.entries.Add(new TableEntry()
                                                {
                                                    name = header,
                                                    value = Util.GetConfigTranslation(config, data[d].ToString(), config.language)
                                                });
                                            }
                                        }

                                    }
                                }
                            }

                            // skip further iteration of excel worksheets
                            break;
                        }
                    }
                }
            }
            sw.Stop();
            long processingTime = sw.ElapsedMilliseconds;
            Console.WriteLine("Excel Attribute gelesen in " + processingTime + "ms.");

            return contents.ToArray();
        }

        /**
         * Get the key of the translation table
         */
        private String GetKey(System.Data.DataTableCollection tables, string category, string subcategory, string fieldname, Languages language, bool description)
        {
            return ExtractI18NAttribute(tables, category, subcategory, fieldname, 7, description);
        }

        /**
            * Get the translation of an attribute
            */
        private string GetTranslation(System.Data.DataTableCollection tables, string category, string subcategory, string fieldname, Languages language, bool description)
        {
            int lng = Languages.DE.Equals(language) ? 9 : 8;
            return ExtractI18NAttribute(tables, category, subcategory, fieldname, lng, description);
        }

        /**
            * Get the translation of an attribute
            */
        private string ExtractI18NAttribute(System.Data.DataTableCollection tables, string category, string subcategory, string fieldname, int column, bool description)
        {
            string translation = "";
            foreach (System.Data.DataTable tableResult in tables)
            {
                // switch the excel tab to i18n
                if (tableResult.TableName == config.workSheetI18NAttributes)
                {

                    for (int i = 0; i < tableResult.Rows.Count; i++)
                    {
                        object[] data = tableResult.Rows[i].ItemArray;
                        int textLength = data[0].ToString().Length;

                        if (textLength > 0 && data[4].ToString().Equals(category) && data[5].ToString().Equals(subcategory) && data[6].ToString().Equals(fieldname)
                            && (!description && "Name".Equals(data[2].ToString()) || description && "Description".Equals(data[2].ToString())))
                        {
                            translation = data[column] as string;
                            break;
                        }
                    }

                    break;
                }
            }

            if (translation == null)
            {
                translation = "";
            }

            return translation;
        }

        /**
         * Get the header translation from another table
         */
        private string GetHeaderTranslation(System.Data.DataTableCollection tables, string key, int row, int cell, Languages language)
        {
            int lng = Languages.DE.Equals(language) ? 2 : 1;
            string translation = key;
            string convertedKey = GetExcelColumnName(cell) + (row + 1).ToString();
            foreach (System.Data.DataTable tableResult in tables)
            {
                if (tableResult.TableName == config.workSheetI18NHeaders)
                {

                    for (int i = 0; i < tableResult.Rows.Count; i++)
                    {
                        object[] data = tableResult.Rows[i].ItemArray;
                        int textLength = data[0].ToString().Length;

                        if (textLength > 0)
                        {
                            string cellKey = data[0].ToString();
                            if (cellKey == "J11")
                            {
                                cellKey = "I13";
                            }
                            
                            if (cellKey.Equals(convertedKey))
                            {
                                translation = data[lng] as string;
                                break;
                            }
                        }
                    }

                    break;
                }
            }

            return translation;
        }

        private string GetExcelColumnName(int columnNumber)
        {
            const int alphabet = 26;
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;
            int modifier = 0;

            while (dividend > 0)
            {
                modulo = (dividend - modifier) % alphabet;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                if (dividend > alphabet)
                {
                    modifier = 1;
                }
                dividend = (int)((dividend - modulo) / alphabet);
            }

            return columnName;
        }

        /**
         * Get the value list values
         */
        private string GetValueList(System.Data.DataTableCollection tables, string key, Languages language)
        {
            string translation = "";
            int lng = Languages.DE.Equals(language) ? 11 : 10;
            List<string> values = new List<string>();
            foreach (System.Data.DataTable tableResult in tables)
            {
                if (tableResult.TableName == config.workSheetValueLists)
                {

                    for (int i = 0; i < tableResult.Rows.Count; i++)
                    {
                        object[] data = tableResult.Rows[i].ItemArray;
                        int textLength = data[0].ToString().Length;

                        if (textLength > 0 && data[9].ToString().Equals(key))
                        {
                            string subkey = key.Replace("ValueList", "ValueListValue");
                            // iterate down
                            for (int j = i + 1; j < tableResult.Rows.Count; j++)
                            {
                                if (tableResult.Rows[j].ItemArray[9].ToString().StartsWith(subkey))
                                {
                                    values.Add(tableResult.Rows[j].ItemArray[lng] as string);
                                }
                            }
                            break;
                        }
                    }

                    break;
                }
            }
            for (int i = 0; i < values.Count; i++)
            {
                translation += values[i];
                if (i + 1 < values.Count)
                {
                    translation += config.valueListSeparator + " ";
                }
            }
            return translation;
        }

    }
}

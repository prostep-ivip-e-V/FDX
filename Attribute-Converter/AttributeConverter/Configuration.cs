using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using IniParser;
using IniParser.Model;

namespace AttributeConverter
{
    class Configuration
    {
        public static Config CreateConfiguration()
        {
            Config config = null;

            try
            {
                var parser = new FileIniDataParser();
                IniParser.Model.Configuration.IniParserConfiguration parserConfig = parser.Parser.Configuration;
                parserConfig.CommentString = "#";
                parserConfig.SectionStartChar = '[';
                parserConfig.SectionEndChar = ']';

                IniData data = parser.ReadFile("settings.ini");

                config = new Config();

                config.language = LanguagesConverter.convert(data["general"]["language"]);
                config.excelFile = data["general"]["excelfile"];
                config.wordTemplate = data["general"]["wordtemplate"];
                config.outputfile = data["general"]["outputfile"];
                config.skipEmtpyLines = "1".Equals(data["general"]["skip_emtpy_lines"]);
                config.valueListSeparator = data["general"]["valuelist_separator"];
                config.processor = int.Parse(data["general"]["processor"]);

                config.columnStructureLevel = int.Parse(data["columns"]["structurelevel"]) - 1;
                config.columnCategory = int.Parse(data["columns"]["category"]) - 1;
                config.columnSubCategory = int.Parse(data["columns"]["subcategory"]) - 1;
                config.columnAttribute = int.Parse(data["columns"]["attribute"]) - 1;
                config.columnStartBlockComponent = int.Parse(data["columns"]["startblock_component"]) - 1;
                config.columnStartBlockAttribute = int.Parse(data["columns"]["startblock_attribute"]) - 1;
                config.columnStartBlockRules = int.Parse(data["columns"]["startblock_rules"]) - 1;
                config.columnBlockComponentRange = int.Parse(data["columns"]["blockrange_component"]);
                config.outputMarkerRow = int.Parse(data["columns"]["output_marker_row"]) - 1;
                config.columnValueLists = data["columns"]["valuelists"];
                config.columnDescription = data["columns"]["attribute_description"];
                config.columnPath = data["columns"]["attribute_path"];

                config.workSheetAttributes = data["excel"]["worksheet.attributes"];
                config.workSheetI18NAttributes = data["excel"]["worksheet.i18n_attributes"];
                config.workSheetI18NHeaders = data["excel"]["worksheet.i18n_headers"];
                config.workSheetValueLists = data["excel"]["worksheet.valuelists"];

                config.translations = new Dictionary<string, string>();
                KeyDataCollection translations = data["translation"];
                foreach (KeyData item in translations)
                {
                    config.translations.Add(item.KeyName, item.Value);
                }

            } catch (Exception e)
            {
                Console.Error.WriteLine(e.ToString());
            }

            return config;
        }

    }
}

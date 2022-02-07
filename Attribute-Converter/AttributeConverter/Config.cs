using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AttributeConverter
{
    class Config
    {
        public Languages language;
        public int columnCategory;
        public int columnSubCategory;
        public int columnAttribute;
        public int columnStructureLevel;
        public string columnDescription;
        public string columnPath;
        public int processor;

        public int columnStartBlockComponent;
        public int columnStartBlockAttribute;
        public int columnStartBlockRules;
        public int columnBlockComponentRange;

        public int outputMarkerRow;

        public string columnValueLists;
        public string valueListSeparator;

        public string excelFile;
        public string wordTemplate;
        public string outputfile;

        public bool skipEmtpyLines;

        public string workSheetAttributes;
        public string workSheetI18NAttributes;
        public string workSheetI18NHeaders;
        public string workSheetValueLists;

        public Dictionary<string, string> translations;

    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

namespace AttributeConverter
{
    class TableEntry
    {
        public bool header = false;
        public Word.WdBuiltinStyle heading = Word.WdBuiltinStyle.wdStyleNormal;
        public string name;
        public string value;
    }
}

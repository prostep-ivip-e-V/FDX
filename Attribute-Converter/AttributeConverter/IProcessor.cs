using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AttributeConverter
{
    interface IProcessor
    {

        void WriteDocument(Table[] content);
    }
}

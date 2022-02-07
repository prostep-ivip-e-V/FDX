using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AttributeConverter
{
    enum Languages
    {
        EN,
        DE
    }

    class LanguagesConverter
    {
        public static Languages convert(string input)
        {
            if (Languages.DE.ToString().ToLower().Equals(input.ToLower()))
            {
                return Languages.DE;
            }
            else if (Languages.EN.ToString().ToLower().Equals(input.ToLower()))
            {
                return Languages.EN;
            }
            return Languages.DE;
        }
    }
}

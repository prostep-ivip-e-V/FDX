using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AttributeConverter
{
    class Util
    {
        public static string GetConfigTranslation(Config config, string input, Languages language)
        {
            string translation = input;

            if (Languages.EN.Equals(language) && config.translations.ContainsKey(input))
            {
                translation = config.translations[input];
            }

            return translation;
        }
    }
}

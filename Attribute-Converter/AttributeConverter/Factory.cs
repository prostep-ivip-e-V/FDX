using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AttributeConverter
{
    class Factory
    {

        public static IProcessor CreateInstance(Config config)
        {
            IProcessor processor;
            switch (config.processor)
            {
                case 0:
                    processor = new MarkDownProcessor(config);
                    break;
                case 1:
                    processor = new WordProcessor(config);
                    break;
                default:
                    processor = new MarkDownProcessor(config);
                    break;

            }
            return processor;
        }

    }
}

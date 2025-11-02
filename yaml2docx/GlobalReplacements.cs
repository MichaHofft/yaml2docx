using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Yaml2Docx
{
    /// <summary>
    /// Because of lack of better ideas, for some "last editing actions", some global replacements
    /// are defined.
    /// </summary>
    public class GlobalReplacements
    {
        public enum Where { 
            Unknown     = 0x0000,
            ColumnFrom  = 0x0001
        }

        public enum How
        {
            FullMatch,
            PartialMatch,
            Regex
        }

        public class Item
        {
            public Where Where = Where.Unknown;
            public How How = How.FullMatch;
            public string From = "";
            public string To = "";
        }

        public Where ParseWhere(string input)
        {
            input = input.Trim().ToLower();
            var res = Where.Unknown;
            if (input == "columnfrom") res = res | Where.ColumnFrom;
            return res;
        }

        public How ParseHow(string input)
        {
            input = input.Trim().ToLower();
            if (input == "fullmatch") return How.FullMatch;
            if (input == "partialmatch") return How.PartialMatch;
            if (input == "regex") return How.Regex;
            return How.FullMatch;
        }

        protected List<Item> _items = new();

        public int ParseListOfString(IEnumerable<string> input)
        {
            int res = 0;
            foreach (var line in input)
            {
                // use "|" as seperator (unlikely to find it in names)
                var parts = line.Split('|');
                if (parts.Length != 4)
                {
                    res--;
                    continue;
                }

                // good to go
                var i = new Item()
                {
                    Where = ParseWhere(parts[0]),
                    How = ParseHow(parts[1]),
                    From = parts[2],
                    To = parts[3]
                };
                _items.Add(i);
                res++;
            }
            return res;
        }

        public string? CheckReplace(Where where, string? input)
        {
            // access
            string? res = input;
            if (res == null)
                return res;

            // most primitive approach first: always iterate thru
            foreach (var it in _items)
            {
                // where matches?
                if ((it.Where & where) == 0)
                    continue;

                // how?
                if (it.How == How.FullMatch)
                {
                    if (input == it.From)
                    {
                        res = it.To;
                        break;
                    }
                }
            }

            return res;
        }
    }
}

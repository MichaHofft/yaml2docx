using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Yaml2Docx
{
    /// <summary>
    /// Used to read grammars, e.g. in BNF or ISO-EBNF
    /// https://github.com/katef/kgt?tab=readme-ov-file
    /// </summary>
    public class GrammarText
    {
        protected Dictionary<string, GrammarPart> _parts = new();

        public GrammarText(string fn)
        {
            Parse(fn);
        }

        public class GrammarPart
        {
            public string Name = "";
            public List<string> Content = new();
        }

        public bool Parse(string fn)
        {
            // access
            string[] lines;
            try 
            {
                lines = System.IO.File.ReadAllLines(fn, encoding: Encoding.UTF8);
            } catch (Exception ex)
            {
                Console.Error.WriteLine($"Exception when accessing {fn}: {ex.Message}");
                return false;
            }

            // ok, split
            GrammarPart? part = null;
            foreach (var line in lines)
            {
                var match = Regex.Match(line, @"^(\w+)\W");
                if (match.Success)
                {
                    // start a new part
                    part = new() { Name = match.Groups[1].ToString() };
                    _parts.Add(part.Name, part);

                    // add also content line
                    part.Content.Add(line);
                }
                else
                if (line.Trim().Length > 0 && part != null)
                {
                    // content line
                    part.Content.Add(line);
                }
            }

            return _parts.Count > 0;
        }

        public GrammarPart? FindPart(string name)
        {
            if (!_parts.ContainsKey(name))
                return null;
            return _parts[name];
        }

        public IEnumerable<string> ListNames()
        {
            foreach (var part in _parts)
                yield return part.Key;
        }

        public static IEnumerable<string> AssembleParts(IEnumerable<GrammarPart> parts)
        {
            var res = new List<string>();
            foreach(var part in parts)
                res.AddRange(part.Content);
            return res;
        }
    }
}

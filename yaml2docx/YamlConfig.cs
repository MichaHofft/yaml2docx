using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;

namespace Yaml2Docx
{
    /// <summary>
    /// This class contains configuration information
    /// </summary>
    public class YamlConfig
    {
        public class ParameterInfo
        {
            public string Name = "";
            public string Description = "";
            public bool Mandatory = false;
            public string Type = "";
            public string Card = "";

            public void Parse(string all)
            {
                var parts = all.Split('|');
                if (parts.Length > 0)
                    Name = parts[0];
                if (parts.Length > 1)
                    Description = parts[1];
                if (parts.Length > 1)
                    Mandatory = (parts[2].ToLower() == "true");
                if (parts.Length > 3)
                    Type = parts[3];
                if (parts.Length > 4)
                    Card = parts[4];
            }

            public string All
            {
                get
                {
                    return $"{Name}|{Description}|{Mandatory}|{Type}|{Card}";
                }
                set
                {
                    Parse(value);
                }
            }

        }

        public class ParameterInfoList : List<ParameterInfo>
        {
            public int FindIndexByName(string name)
            {
                for (int i = 0; i < this.Count; i++)
                    if (this[i].Name == name)
                        return i;
                return -1;
            }

            public void AddOrReplace(ParameterInfo pi)
            {
                int idx = FindIndexByName(pi.Name);
                if (idx >= 0)
                    this[idx] = pi;
                else
                    this.Add(pi);
            }

            public void AddOrReplace(ParameterInfoList list)
            {
                foreach (var pi in list)
                    AddOrReplace(pi);
            }

            public void RemoveByName(string name)
            {
                int idx = FindIndexByName(name);
                if (idx >= 0)
                    this.RemoveAt(idx);
            }
        }

        public class OperationConfig
        {
            public string? Heading;
            public string? Body;

            public string? Explanation;

            public ParameterInfoList Inputs = new();
            public ParameterInfoList Outputs = new();

            public List<string> SuppressInputs = new();
            public List<string> SuppressOutputs = new();
        }

        public class ReadOpenApiFile
        {
            public string Fn = "TBD.yaml";

            public string? Heading2Text;
            public string? Body2Text;

            public bool ListOperations = true;
            
            public Dictionary<string, OperationConfig> ExportOperations = new();
        }

        public class CreateWordFile
        {
            public string Fn = "TBD.docx";
            public string? UseTemplateFn;

            public List<ReadOpenApiFile> ReadOpenApiFiles = new();
        }

        public class ExportConfig
        {
            public string Heading3 = "TBD";
            public string Body = "TBD";

            public string Heading2Style = "Normal";
            public string Heading3Style = "Normal";
            public string BodyStyle = "Normal";
            public string TableCaptionStyle = "Normal";

            public List<double>? TableColumnWidthCm;

            public bool AddTableCaptions = true;

            public ParameterInfoList Inputs = new();
            public ParameterInfoList Outputs = new();

            public List<string> SuppressInputs = new();
            public List<string> SuppressOutputs = new();
            
            public List<CreateWordFile> CreateWordFiles = new();
        }

        public static ExportConfig Load(string fn)
        {
            var yml = System.IO.File.ReadAllText(fn);
            var deserializer = new DeserializerBuilder()
                // convert YAML "someKey" to C# "SomeKey"
                .WithNamingConvention(CamelCaseNamingConvention.Instance)
                // .IgnoreUnmatchedProperties()
                .Build();
            var cfg = deserializer.Deserialize<ExportConfig>(yml);
            return cfg;
        }
    }
}

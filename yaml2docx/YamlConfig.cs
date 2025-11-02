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

            public List<string>? Notes;

            public ParameterInfoList Inputs = new();
            public ParameterInfoList Outputs = new();

            public List<string> SuppressInputs = new();
            public List<string> SuppressOutputs = new();
        }

        public class ExportAction
        {
            /// <summary>
            /// ExportPara, ExportOverview, ExportTables, ExportYaml, ExportSchemas, ExportPatterns,
            /// ExportRailRoad, ExportGrammar
            /// </summary>
            public string Action = "";

            public string? ParaText;
            public string? ParaStyle;

            public bool YamlAsSource = false;
            public bool YamlAsTable = false;

            public bool SkipIfVisited = false;

            public List<string>? SchemaNotFollow;

            public List<string> IncludeSchemas = new();
            public List<string> SuppressSchemas = new();
            public List<string> SuppressMembers = new();

            public List<string> Parts = new();

            /// <summary>
            /// Console Svg Utf8
            /// </summary>
            public string OutputFormat = "";

            public string? Heading;
            public string? Body;
            public List<string>? Notes;

            public double? FontSize = null;
            public double? TargetWidthCm = null;
            public double? CropBottomCm = null;
        }

        public class ReadOpenApiFile
        {
            public bool Skip = false;

            public string Fn = "TBD.yaml";

            public bool ListOperations = false;

            public List<ExportAction> Actions = new();
            
            public Dictionary<string, OperationConfig> UseOperations = new();
        }

        public class ReadRailRoadFile
        {
            public bool Skip = false;

            public string Fn = "TBD.txt";

            public bool ListNames = false;

            public List<ExportAction> Actions = new();
        }

        public class ReadGrammarFile
        {
            public bool Skip = false;

            public string Fn = "TBD.txt";

            public bool ListNames = false;

            public List<ExportAction> Actions = new();
        }

        public class CreateWordFile
        {
            public string Fn = "TBD.docx";
            public string? UseTemplateFn;

            public bool ListStyles = false;

            public List<ReadOpenApiFile> ReadOpenApiFiles = new();
            public List<ReadRailRoadFile> ReadRailRoadFiles = new();
            public List<ReadGrammarFile> ReadGrammarFiles = new();
        }

        public class ExportConfig
        {
            public string TableHeadingPrefix = "TBD";
            public string Body = "TBD";

            public string SchemaHeadingPrefix = "TBD";
            public string SchemaBody = "TBD";

            public string SchemaTableCaptionPrefix = "TBD";
            public string PatternTableCaptionPrefix = "TBD";

            public string Heading2Style = "Normal";
            public string TableHeadingStyle = "Normal";
            public string SchemaHeadingStyle = "Normal";
            public string BodyStyle = "Normal";
            public string NoteStyle = "Normal";
            public string TableCaptionStyle = "Normal";
            public string YamlHeadingStyle = "Normal";
            public string YamlCodeStyle = "Normal";
            public string GrammarHeadingStyle = "Normal";
            public string GrammarCodeStyle = "Normal";

            public double? GrammarCodeFontSize = null;

            public double? GrammarCodeTargetWidthCm = 16.0;
            public double? GrammarCodeMaxHeightCm = 22.0;

            public string DockerBuildTextCmd = "docker";
            public string DockerBuildTextArgs = "run --rm -i -v \".:/data\" kgt -l iso-ebnf -e rrutf8";

            public string DockerBuildSvgCmd = "docker";
            public string DockerBuildSvgArgs = "run --rm -i -v \".:/data\" kgt -l iso-ebnf -e svg";

            public string DockerSvg2BitmapCmd = "docker";
            public string DockerSvgBitmapfArgs = "run --rm -v \"%wd%:/data\" -w /data homi/librsvg --background-color=white --width=4000px -f png -o \"%out-fn%\" \"%in-fn%\"";

            public uint TableCellBorderWidth = 8;
            public uint YamlMonoBorderWidth = 8;

            public int NumberEmptyLines = 1;

            public List<double>? TableColumnWidthCm;
            public List<double>? OverviewColumnWidthCm;
            public List<double>? InterfaceOpFiveColumnWidthCm;
            public List<double>? InterfaceOpThreeColumnWidthCm;

            public bool AddTableCaptions = true;

            public ParameterInfoList Inputs = new();
            public ParameterInfoList Outputs = new();

            public List<string> SuppressInputs = new();
            public List<string> SuppressOutputs = new();

            public List<string> OriginSchemaOrder = new();
            public List<string> SuppressSchemaNames = new();

            public int PatternInlineLimit = 80;

            public List<string> GlobalReplacements = new();

            [YamlIgnore]
            public GlobalReplacements Reps = new();

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
            cfg.Reps.ParseListOfString(cfg.GlobalReplacements);
            return cfg;
        }
    }
}

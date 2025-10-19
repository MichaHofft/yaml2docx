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
    /// This class provides a set of C# classes to represent an OpenAPI document in YAML format.
    /// Warning: This is a simplified representation and may not cover all aspects of the OpenAPI specification.
    ///          These classes were reversely engineered from a specific OpenAPI YAML file and may need adjustments 
    ///          for other files. Particularily, they even may be wrong and not fully compliant to the OpenAPI spec!
    /// </summary>
    public class YamlOpenApi
    {
        public class OpenApiInfo
        {
            public string? Title;
            public string? Description;

            public Dictionary<string, string>? Contact;
            public Dictionary<string, string>? License;

            public string? Version;

            [YamlDotNet.Serialization.YamlMember(Alias = "x-profile-identifier", ApplyNamingConventions = false)]
            public string ProfileIdentifier = "";
        }

        public class OpenApiServer
        {
            public string Url { get; set; } = string.Empty;
        }

        public class OpenApiProperty
        {
            public string? Type;
            public string? Format;
            public string? Pattern;
            public string? Description;
            public string? Example;

            public int minItems = 0;
            public int maxItems = 0;
            public int minLength = 0;
            public int maxLength = 0;

            public List<string>? Enum;

            public OpenApiItems? Items;

            public List<OpenApiProperty>? AllOf;

            [YamlDotNet.Serialization.YamlMember(Alias = "$ref", ApplyNamingConventions = false)]
            public string? Ref;
        }

        public class OpenApiItems
        {
            public string? Type;
            public string? Example;
            
            [YamlDotNet.Serialization.YamlMember(Alias = "$ref", ApplyNamingConventions = false)]
            public string? Ref;
        }

        public class OpenApiSchemaPart
        {
            // Note: this might indicated, that the part class needs to be restructured!!
            public string? Type;
            
            // list of required property names
            public List<string>? Required;

            // define further properties
            public Dictionary<string, OpenApiProperty>? Properties;

            // just refer to other schemas to be part of this schema
            [YamlDotNet.Serialization.YamlMember(Alias = "$ref", ApplyNamingConventions = false)]
            public string? Ref;
        }

        public class OpenApiSchema
        {
            public string? Description;
            public string? Type;
            public string? Format;
            public string? Default;
            public string? Minimum;
            public string? Maximum;

            public string? Pattern;
            public string? Example;

            public Dictionary<string, OpenApiProperty>? Properties;
            
            public OpenApiItems? Items;

            public List<string>? Enum;

            public List<OpenApiSchemaPart>? AllOf;

            public List<OpenApiSchemaPart>? OneOf;

            // list of required property names
            public List<string>? Required;

            // reference to another schema
            [YamlDotNet.Serialization.YamlMember(Alias = "$ref", ApplyNamingConventions = false)]
            public string? Ref;
        }

        public class OpenApiEncoding
        {
            public string? ContentType;
        }

        public class OpenApiContent
        {
            public OpenApiSchema? Schema;
            public Dictionary<string, OpenApiEncoding> Encoding = new();
        }

        public class OpenApiHeader
        {
            public string? Description;
            public string? Example;
            
            public bool Required = false;
            
            public OpenApiSchema? Schema;

            // just a reference to another header
            [YamlDotNet.Serialization.YamlMember(Alias = "$ref", ApplyNamingConventions = false)]
            public string? Ref;
        }

        public class OpenApiResponse
        {
            public string? Description;

            public Dictionary<string, OpenApiHeader> Headers = new();

            public Dictionary<string, OpenApiContent> Content = new();

            // just a reference to another response
            [YamlDotNet.Serialization.YamlMember(Alias = "$ref", ApplyNamingConventions = false)]
            public string? Ref;
        }

        public class OpenApiRequestBody
        {
            public string? Description;
            
            public bool Required = false;
            
            public Dictionary<string, OpenApiContent> Content = new();
        }

        public class OpenApiParameter
        {
            public string? Name;
            public string? In;
            public string? Description;
            public string? Style;

            public bool Required = false;
            public bool Deprecated = false;
            public bool Explode = false;

            public OpenApiSchema? Schema;

            // just a reference to another parameter
            [YamlDotNet.Serialization.YamlMember(Alias = "$ref", ApplyNamingConventions = false)]
            public string? Ref;
        }

        public class OpenApiOperation
        {
            public string? Summary;
            public string? OperationId;

            public List<string> Tags = new();
            
            [YamlMember(Alias = "x-semanticIds", ApplyNamingConventions = false)]
            public List<string>? SemanticIds;

            public List<OpenApiParameter>? Parameters;

            public OpenApiRequestBody? RequestBody;

            public Dictionary<string, OpenApiResponse>? Responses;
        }

        public class OpenApiPath
        {
            public List<OpenApiParameter>? Parameters;

            public OpenApiOperation? Get ;
            public OpenApiOperation? Put ;
            public OpenApiOperation? Patch ;
            public OpenApiOperation? Post ;
            public OpenApiOperation? Delete ;

            public IEnumerable<OpenApiOperation> Operations()
            {
                if (Get != null) yield return Get;
                if (Put != null) yield return Put;
                if (Patch != null) yield return Patch;
                if (Post != null) yield return Post;
                if (Delete != null) yield return Delete;
            }

            public IEnumerable<Tuple<string, OpenApiOperation>> FullOperations()
            {
                if (Get != null) yield return new Tuple<string, OpenApiOperation>("Get", Get);
                if (Put != null) yield return new Tuple<string, OpenApiOperation>("Put", Put);
                if (Patch != null) yield return new Tuple<string, OpenApiOperation>("Patch", Patch);
                if (Post != null) yield return new Tuple<string, OpenApiOperation>("Post", Post);
                if (Delete != null) yield return new Tuple<string, OpenApiOperation>("Delete", Delete);
            }
        }

        public class OpenApiComponents
        {
            public Dictionary<string, OpenApiSchema>? Schemas;
            public Dictionary<string, OpenApiResponse>? Responses;
            public Dictionary<string, OpenApiParameter>? Parameters;
            public Dictionary<string, OpenApiRequestBody>? RequestBodies;
        }

        public class OpenApiDocument
        {
            [YamlMember(Alias = "openapi", ApplyNamingConventions = false)]
            public string? OpenApiVersion;

            public OpenApiInfo? Info;
            public OpenApiComponents? Components;

            public List<OpenApiServer>? Servers;
            public Dictionary<string, OpenApiPath>? Paths;

            public OpenApiOperation? FindApiOperation(string operationId)
            {
                if (Paths == null) 
                    return null;
                foreach (var pathEntry in Paths)
                {
                    var path = pathEntry.Value;
                    foreach (var operation in path.Operations())
                    {
                        if (operation.OperationId == operationId)
                            return operation;
                    }
                }
                return null;
            }
        }

        public static OpenApiDocument Load(string fn)
        {
            var yml = System.IO.File.ReadAllText(fn);
            var deserializer = new DeserializerBuilder()
                // convert YAML "someKey" to C# "SomeKey"
                .WithNamingConvention(CamelCaseNamingConvention.Instance)
                // ---
                // allow only properties defined in the C# classes, therefore disable this:
                // .IgnoreUnmatchedProperties()
                // ---
                // Not required/ useful/ meaningful:
                // .WithNamingConvention(UnderscoredNamingConvention.Instance)
                // .WithNamingConvention(NullNamingConvention.Instance)
                // ---
                .Build();
            var doc = deserializer.Deserialize<OpenApiDocument>(yml);
            return doc;
        }
    }
}

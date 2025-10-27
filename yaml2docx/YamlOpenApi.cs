using System;
using System.Collections.Generic;
using System.Globalization;
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

        /// <summary>
        /// Purpose: Describe the structure of data objects
        /// </summary>
        public class OpenApiProperty
        {
            public string? Type;
            public string? Format;
            public string? Pattern;
            public string? Description;
            public string? Example;

            public int? minItems;
            public int? maxItems;
            public int? minLength;
            public int? maxLength;

            public List<string>? Enum;

            public OpenApiItems? Items;

            public List<OpenApiProperty>? AllOf;

            [YamlDotNet.Serialization.YamlMember(Alias = "$ref", ApplyNamingConventions = false)]
            public string? Ref;

            public OpenApiProperty Clone()
            {
                var res = new OpenApiProperty()
                {
                    Type = Type,
                    Format = Format,
                    Pattern = Pattern,
                    Description = Description,
                    Example = Example,
                    minItems = minItems,
                    maxItems = maxItems,
                    minLength = minLength,
                    maxLength = maxLength,
                    Ref = Ref
                };

                if (Enum != null)
                    res.Enum = new List<string>(Enum);

                if (Items != null)
                    res.Items = Items.Clone();

                if (AllOf != null)
                    res.AllOf = new List<OpenApiProperty>(AllOf);

                return res;
            }

            public void Join(OpenApiProperty other)
            {
                if (Type == null)
                    Type = other.Type;
                if (Format == null)
                    Format = other.Format;
                if (Pattern == null)
                    Pattern = other.Pattern;
                if (Description == null)
                    Description = other.Description;
                if (Example == null)
                    Example = other.Example;

                if (minItems == null)
                    minItems = other.minItems;
                if (maxItems == null)
                    maxItems = other.maxItems;
                if (minLength == null)
                    minLength = other.minLength;
                if (maxLength == null)
                    maxLength = other.maxLength;

                if (Enum == null) 
                    Enum = other.Enum;
                else if (other.Enum != null)
                    Enum.AddRange(other.Enum);

                if (Items == null)
                    Items = other.Items;

                if (AllOf == null)
                    AllOf = other.AllOf;
                else if (other.AllOf != null)
                    AllOf.AddRange(other.AllOf);

                // don't knwo that to do with Ref ..
                if (Ref != null && other.Ref != null) 
                    throw new Exception("Cannot handle joining to Refs!");
            }

            public void SetFrom(OpenApiSchema other)
            {
                if (other.Type != null)
                    Type = other.Type;
                if (other.Enum != null)
                    Enum = other.Enum;
            }
        }

        public class OpenApiOriginatedProperty
        {
            /// <summary>
            /// Component coming from
            /// </summary>
            public string? Origin;

            /// <summary>
            /// Name, as given by dictionary key
            /// </summary>
            public string? Name;

            /// Requested as required by superior data structure            
            public bool Required;

            /// <summary>
            /// Property itself, value of dictionary
            /// </summary>
            public OpenApiProperty Property = new();

            public OpenApiOriginatedProperty(string origin, string name, bool required, OpenApiProperty property)
            {
                Origin = origin;
                Name = name;
                Required = required;
                Property = property;
            }
        }

        public class OpenApiOriginatedPropertyList : List<OpenApiOriginatedProperty>
        {
            public OpenApiOriginatedPropertyList() : base()
            {
            }

            public OpenApiOriginatedPropertyList(IEnumerable<OpenApiOriginatedProperty> items) : base(items)
            {
            }
        }

        public class OpenApiItems
        {
            public string? Type;
            public string? Example;
            
            [YamlDotNet.Serialization.YamlMember(Alias = "$ref", ApplyNamingConventions = false)]
            public string? Ref;

            public OpenApiItems Clone()
            {
                var res = new OpenApiItems()
                {
                    Type = Type,
                    Example = Example,
                    Ref = Ref
                };
                return res;
            }
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

            public int? MinLength;
            public int? MaxLength;

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

            public OpenApiResponse Clone()
            {
                var res = new OpenApiResponse();
                
                res.Description = Description;
                res.Headers = new Dictionary<string, OpenApiHeader>(Headers);
                res.Content = new Dictionary<string, OpenApiContent>(Content);

                return res;
            }

            public void Join(OpenApiResponse? other)
            {
                if (other == null)
                    return;

                if (Description == null)
                    Description = other.Description;

                if (other.Headers != null)
                {
                    Headers = Headers ?? new Dictionary<string, OpenApiHeader>();
                    foreach (var x in other.Headers)
                        Headers.Add(x.Key, x.Value);
                }

                if (other.Content != null)
                {
                    Content = Content ?? new Dictionary<string, OpenApiContent>();
                    foreach (var x in other.Content)
                        Content.Add(x.Key, x.Value);
                }
            }
        }

        public class OpenApiRequestBody
        {
            public string? Description;
            
            public bool Required = false;
            
            public Dictionary<string, OpenApiContent> Content = new();
        }

        /// <summary>
        /// Purpose: Describe inputs to an API endpoint
        /// </summary>
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

            public object? FindComponent(string path)
            {
                Func<string, string?> check = (start) =>
                {
                    if (path.StartsWith(start))
                        return path.Substring(start.Length);
                    return null;
                };

                var isSchema = check("#/components/schemas/");
                if (isSchema != null && true == Components?.Schemas?.ContainsKey(isSchema))
                    return Components?.Schemas[isSchema];

                var isResponses = check("#/components/responses/");
                if (isResponses != null && true == Components?.Responses?.ContainsKey(isResponses))
                    return Components?.Responses[isResponses];

                var isParameters = check("#/components/parameters/");
                if (isParameters != null && true == Components?.Parameters?.ContainsKey(isParameters))
                    return Components?.Parameters[isParameters];

                var isRequestBodies = check("#/components/requestBodies/");
                if (isRequestBodies != null && true == Components?.RequestBodies?.ContainsKey(isRequestBodies))
                    return Components?.RequestBodies[isRequestBodies];

                return null;
            }

            public T? FindComponent<T>(string path) where T : class
            {
                return FindComponent(path) as T;
            }

            public OpenApiOriginatedPropertyList? RecursiveFindPropertyBundles(
                string schemaName,
                HashSet<string>? schemasTouched = null)
            {
                // any result?
                var schema = FindComponent<OpenApiSchema>(schemaName);
                if (schema == null) 
                    return null;

                if (schemaName == "HasKind")
                    ;

                // ok, start result
                var res = new OpenApiOriginatedPropertyList();

                // some lambdas for touching schemas
                Action<string?> touchSchema = (sname) =>
                {
                    if (sname != null)
                        schemasTouched?.Add(sname);
                };

                Action<string?> expandAndTouchSchema = (sname) =>
                {
                    // access
                    if (sname == null)
                        return;

                    // potential component?
                    var typeComp = FindComponent<YamlOpenApi.OpenApiSchema>("#/components/schemas/" + sname);
                    // is a one of?
                    if (typeComp != null && typeComp.OneOf != null && typeComp.OneOf.Count > 0)
                    {
                        foreach (var one in typeComp.OneOf)
                            touchSchema(YamlOpenApi.StripSchemaHead(one?.Ref));
                    }
                    else
                        touchSchema(sname);
                };

                // integrate all allOf
                if (schema.AllOf != null)
                    foreach (var ao in schema.AllOf)
                    {
                        // do we find properties
                        if (ao?.Properties?.Any() == true)
                            foreach (var p in ao.Properties)
                            {
                                if (p.Key == "kind")
                                    ;

                                // make a new list of property attributes
                                var joinedProp = p.Value.Clone();

                                // may be SET by a Ref
                                if (p.Value.Ref != null)
                                {
                                    var refSchema = FindComponent<OpenApiSchema>(p.Value.Ref);
                                    if (refSchema != null)
                                    {
                                        // overtake attributes!
                                        joinedProp.SetFrom(refSchema);
                                    }
                                }

                                // for properties.AllOf -> join ATTRIBUTES together ..
                                if (p.Value.AllOf != null)
                                    foreach (var poa in p.Value.AllOf)
                                        joinedProp.Join(poa);

                                // use type to add to touched schemas
                                touchSchema(p.Value.Type);
                                expandAndTouchSchema(YamlOpenApi.StripSchemaHead(p.Value.Ref));
                                if (p.Value.Type == "array" && p.Value.Items?.Ref != null)
                                    expandAndTouchSchema(YamlOpenApi.StripSchemaHead(p.Value.Items.Ref));

                                // add property
                                res.Add(new OpenApiOriginatedProperty(schemaName, p.Key, IsContained(ao.Required, p.Key), joinedProp));
                            }

                        // TODO: what is with properties.Items / AllOf

                        // do we find references
                        if (ao?.Ref != null)
                        {
                            // add to touched
                            var refSchemaName = StripSchemaHead(ao.Ref);
                            touchSchema(refSchemaName);

                            // recurse
                            var props2 = RecursiveFindPropertyBundles(ao.Ref, schemasTouched);
                            if (props2 != null)
                                res.AddRange(props2);
                        }
                    }

                // also "normal" properties
                if (schema.Properties?.Any() == true)
                    foreach (var sp in schema.Properties)
                    {
                        // for Ref of property -> join ATTRIBUTES together ..
                        var joinedProp = sp.Value.Clone();
                        if (sp.Value.Ref != null)
                        {
                            var refSchema = FindComponent<OpenApiSchema>(sp.Value.Ref);
                            if (refSchema != null)
                            {
                                // overtake attributes!
                                joinedProp.SetFrom(refSchema);
                            }
                        }

                        // use type to add to touched schemas
                        expandAndTouchSchema(joinedProp.Type);
                        if (joinedProp.Type == "array" && joinedProp.Items?.Ref != null)
                            expandAndTouchSchema(YamlOpenApi.StripSchemaHead(joinedProp.Items.Ref));

                        // add property
                        res.Add(new OpenApiOriginatedProperty(schemaName, sp.Key, IsContained(schema.Required, sp.Key), joinedProp));
                    }

                // in any case a success
                return res;
            }            
        }

        public static OpenApiDocument? Load(string fn)
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

        public static string? StripSchemaHead(string? refStr)
        {
            return refStr?.Replace("#/components/schemas/", "");
        }

        public static string? StripResponseHead(string? refStr)
        {
            return refStr?.Replace("#/components/responses/", "");
        }

        public static bool IsContained(List<string>? list, string? val)
        {
            if (list == null || val == null)
                return false;
            foreach (var l in list)
                if (l.Equals(val, StringComparison.InvariantCultureIgnoreCase))
                    return true;
            return false;
        }

    }
}

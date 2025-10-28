using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Extensions;
using YamlDotNet.Core;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;
using YamlDotNet.Serialization.ObjectGraphVisitors;
using static Yaml2Docx.YamlOpenApi;

namespace Yaml2Docx
{
    /// <summary>
    /// This class exports small tables for the description on Interface Operations in IEC 63278-5.
    /// </summary>
    public class ExportIecInterfaceOperation
    {
        protected YamlConfig.ExportConfig _config = new YamlConfig.ExportConfig();

        public ExportIecInterfaceOperation(YamlConfig.ExportConfig config)
        {
            _config = config;
        }

        public void ExportParagraph(
            MainDocumentPart mainPart,
            string? text,
            string? style)
        {
            // access
            Body? body = mainPart.Document.Body;
            if (body == null)
                return;

            // generate a table-reference to the last table WITHOUT creating NEW TABLE
            var substs = new List<Substitution>() {
                new Substitution("table-ref", $"Table{_tableRefIdCount}", isBookmark: true)
            };

            // ok
            if (text != null)
                body.AppendChild(CreateParagraph(
                    $"{text}",
                    styleId: $"{style ?? "Normal"}",
                    substitutions: substs));
        }

        /// <summary>
        /// Substitution for doing variable replacement in paragraphs
        /// </summary>
        public record Substitution(string Key, string Value, bool isBookmark = false);

        protected static int _tableRefIdCount = 13;

        public record OperationTuple (YamlConfig.OperationConfig Config, YamlOpenApi.OpenApiOperation Operation);

        /// <summary>
        /// Export a single operation to the Word
        /// </summary>
        public void ExportSingleOperation(
            MainDocumentPart mainPart,
            YamlConfig.OperationConfig opConfig,
            YamlOpenApi.OpenApiOperation op)
        {
            // access
            Body? body = mainPart.Document.Body;
            if (body == null)
                return;

            // generate a table-reference
            var substTablRef = new Substitution("table-ref", $"Table{_tableRefIdCount++}", isBookmark: true);
            var substs = new List<Substitution>() { substTablRef };

            // Heading
            body.AppendChild(CreateParagraph(
                $"{opConfig?.Heading ?? _config.TableHeadingPrefix} {op.OperationId}",
                styleId: $"{_config.TableHeadingStyle}"));

            // Intro text
            body.AppendChild(CreateParagraph(
                $"{opConfig?.Body ?? _config.Body}",
                styleId: $"{_config.BodyStyle}",
                substitutions: substs));

            // Create input, output parameters
            YamlConfig.ParameterInfoList inputs = new();
            if (_config.Inputs != null)
                inputs.AddRange(_config.Inputs);
            if (opConfig?.Inputs != null)
                inputs.AddOrReplace(opConfig.Inputs);

            YamlConfig.ParameterInfoList outputs = new();
            if (_config.Outputs != null)
                outputs.AddRange(_config.Outputs);
            if (opConfig?.Outputs != null)
                outputs.AddOrReplace(opConfig.Outputs);

            // turn the operation's parameters into inputs
            if (op.Parameters != null)
                foreach (var param in op.Parameters)
                {
                    var pi = new YamlConfig.ParameterInfo()
                    {
                        Name = param.Name ?? "",
                        Description = param.Description ?? "",
                        Mandatory = (param.Required == true) ? true : false,
                        Type = param.Schema?.Type ?? "TBD",
                        Card = "1"
                    };
                    if (pi.Mandatory == false)
                        pi.Card = "0..1";
                    inputs.AddOrReplace(pi);
                }

            // ok, suppress, BUT EXCLUSIVE in groups
            if (opConfig?.SuppressInputs != null)
            {
                foreach (var name in opConfig.SuppressInputs)
                    inputs.RemoveByName(name);
            }
            else
            if (_config.SuppressInputs != null)
            {
                foreach (var name in _config.SuppressInputs)
                    inputs.RemoveByName(name);
            }

            if (_config.SuppressOutputs != null)
            {
                foreach (var name in _config.SuppressOutputs)
                    inputs.RemoveByName(name);
            }
            else
            if (_config.SuppressOutputs != null)
            {
                foreach (var name in _config.SuppressOutputs)
                    inputs.RemoveByName(name);
            }                      

            // build explanation
            var explanation = opConfig?.Explanation ?? op?.Summary;

            // Create the table
            Table table = new Table();

            // Define table properties (1 pt border, full width)
            TableProperties tblProps = new TableProperties(
                new TableWidth { Type = TableWidthUnitValues.Pct, Width = "5000" }, // 100% width (5000 = 100% in OpenXML)
                new TableLayout { Type = TableLayoutValues.Fixed }, // <=== FIXED LAYOUT
                new TableBorders(
                    new TopBorder { Val = BorderValues.Single, Size = 8 },
                    new BottomBorder { Val = BorderValues.Single, Size = 8 },
                    new LeftBorder { Val = BorderValues.Single, Size = 8 },
                    new RightBorder { Val = BorderValues.Single, Size = 8 },
                    new InsideHorizontalBorder { Val = BorderValues.Single, Size = 8 },
                    new InsideVerticalBorder { Val = BorderValues.Single, Size = 8 }
                )
            );
            table.AppendChild(tblProps);

            // In OpenXML, table and cell widths are controlled by:
            // TableProperties → TableWidth controls the overall table width.
            // TableGrid → defines each column’s width explicitly.
            // TableCellProperties → optional, but can reinforce column widths.
            // 
            // All widths are in twips (1/20th of a point), where
            // 1 inch = 1440 twips, 1 cm = 567 twips.
            // Standard Word page width ≈ 12240 twips (8.5 inches)
            // Page margins usually take ~1 inch each side, so usable width ≈ 9360 twips.
            // 
            // So, if you want the table to fill the text width, we can use that 9360 twip range.^

            // Define column widths (sum ~9360 twips = ~6.5 inches)
            double cm = 567;
            int[] cw = { (int)(3 * cm), (int)(6 * cm), (int)(1 * cm), (int)(4 * cm), (int)(1 * cm) };

            if (_config.TableColumnWidthCm != null && _config.TableColumnWidthCm.Count >= 5)
                for (int i = 0; i < Math.Min(5, _config.TableColumnWidthCm.Count); i++)
                    cw[i] = (int)(cm * _config.TableColumnWidthCm[i]);

            TableGrid tableGrid = new TableGrid();
            foreach (int width in cw)
            {
                tableGrid.Append(new GridColumn() { Width = width.ToString() });
            }
            table.Append(tableGrid);

            // special col widths
            int cw04 = cw[0] + cw[1] + cw[2] + cw[3] + cw[4];
            int cw14 = cw[1] + cw[2] + cw[3] + cw[4];

            // 1st row: Header for interface operation
            {
                TableRow tr = new TableRow();
                tr.Append(CreateCell("Interface Operation Name ", cw[0]));
                tr.Append(CreateMergedCell($"{op?.OperationId}", true, cw[1], bold: true));
                tr.Append(CreateMergedCell("", false, cw[2]));
                tr.Append(CreateMergedCell("", false, cw[3]));
                tr.Append(CreateMergedCell("", false, cw[4]));
                table.Append(tr);
            }

            // 2nd row: Explanation
            {
                TableRow tr = new TableRow();
                tr.Append(CreateCell("Explanation", cw[0]));
                tr.Append(CreateMergedCell($"{explanation}", true, cw[1]));
                tr.Append(CreateMergedCell("", false, cw[2]));
                tr.Append(CreateMergedCell("", false, cw[3]));
                tr.Append(CreateMergedCell("", false, cw[4]));
                table.Append(tr);
            }

            //
            // Input
            //

            // 3rd row: Input
            {
                TableRow tr = new TableRow();
                tr.Append(CreateCell("Name", cw[0]));
                tr.Append(CreateCell("Description", cw[1]));
                tr.Append(CreateCell("Mand.", cw[2]));
                tr.Append(CreateCell("Type", cw[3]));
                tr.Append(CreateCell("Card.", cw[4]));
                table.Append(tr);
            }

            // 4th row: Input parameters
            if (inputs.Count > 0)
            {
                TableRow tr = new TableRow();
                tr.Append(CreateMergedCell("Input parameter(s)", true, cw04));
                tr.Append(CreateMergedCell("", false, cw[1]));
                tr.Append(CreateMergedCell("", false, cw[2]));
                tr.Append(CreateMergedCell("", false, cw[3]));
                tr.Append(CreateMergedCell("", false, cw[4]));
                table.Append(tr);
            }

            // 5th.. row: Single input parameter
            foreach (var pi in inputs)
            {
                TableRow tr = new TableRow();
                tr.Append(CreateCell(pi.Name, cw[0]));
                tr.Append(CreateCell(pi.Description, cw[1]));
                tr.Append(CreateCell(pi.Mandatory ? "yes" : "no", cw[2]));
                tr.Append(CreateCell(pi.Type, cw[3]));
                tr.Append(CreateCell(pi.Card, cw[4]));
                table.Append(tr);
            }

            //
            // Request body .. treated as input
            //

            if (op?.RequestBody?.Content != null && op?.RequestBody.Content.Count > 0)
            {
                // try to compile appropriate information
                var name = "requestBody";
                var desc = op.RequestBody.Description ?? "\u2014";
                var mand = $"{(op.RequestBody.Required ? "yes" : "no")}";
                var type = "\u2014";
                var card = $"{(op.RequestBody.Required ? "1" : "0..1")}";

                // Invisible to the reader: multiple content types/ schemas, take the first as type
                foreach (var cntTup in op.RequestBody.Content)
                    if (cntTup.Value?.Schema?.Ref != null)
                    {
                        type = YamlOpenApi.StripSchemaHead(cntTup.Value.Schema.Ref);
                        break;
                    }

                // do it
                TableRow tr = new TableRow();
                tr.Append(CreateCell($"{name}", cw[0]));
                tr.Append(CreateCell($"{desc}", cw[1]));
                tr.Append(CreateCell($"{mand}", cw[2]));
                tr.Append(CreateCell($"{type}", cw[3]));
                tr.Append(CreateCell($"{card}", cw[4]));
                table.Append(tr);
            }

            //
            // Output
            //

            // 6th row: Output parameters
            if (outputs.Count > 0)
            {
                TableRow tr = new TableRow();
                tr.Append(CreateMergedCell("Output parameter(s)", true, cw04));
                tr.Append(CreateMergedCell("", false, cw[1]));
                tr.Append(CreateMergedCell("", false, cw[2]));
                tr.Append(CreateMergedCell("", false, cw[3]));
                tr.Append(CreateMergedCell("", false, cw[4]));
                table.Append(tr);
            }

            // 7th.. row: Single output parameter
            foreach (var pi in outputs)
            {
                TableRow tr = new TableRow();
                tr.Append(CreateCell(pi.Name, cw[0]));
                tr.Append(CreateCell(pi.Description, cw[1]));
                tr.Append(CreateCell(pi.Mandatory ? "yes" : "no", cw[2]));
                tr.Append(CreateCell(pi.Type, cw[3]));
                tr.Append(CreateCell(pi.Card, cw[4]));
                table.Append(tr);
            }

            //
            // Response body .. treated as output
            //

            if (op?.Responses != null && op?.Responses.Count > 0)
                foreach (var resp in op.Responses)
                {
                    // filter out the 200 range
                    if (resp.Key?.StartsWith("2") != true || resp.Value == null)
                        continue;

                    // try to compile appropriate information
                    var name = "responseBody";
                    var desc = resp.Value.Description ?? "\u2014";
                    var mand = "yes";
                    var type = "\u2014";
                    var card = "1";

                    // Invisible to the reader: multiple content types/ schemas, take the first as type
                    if (resp.Value.Content != null)
                        foreach (var cntTup in resp.Value.Content)
                            if (cntTup.Value?.Schema?.Ref != null)
                            {
                                type = YamlOpenApi.StripSchemaHead(cntTup.Value.Schema.Ref);
                                break;
                            }

                    // do it
                    TableRow tr = new TableRow();
                    tr.Append(CreateCell($"{name}", cw[0]));
                    tr.Append(CreateCell($"{desc}", cw[1]));
                    tr.Append(CreateCell($"{mand}", cw[2]));
                    tr.Append(CreateCell($"{type}", cw[3]));
                    tr.Append(CreateCell($"{card}", cw[4]));
                    table.Append(tr);
                }

            //
            // Caption
            //

            // Before appending the table, add some caption text?
            if (_config.AddTableCaptions)
            {
                // Caption paragraph
                Paragraph caption = new Paragraph(
                    new ParagraphProperties(
                        new ParagraphStyleId { Val = _config.TableCaptionStyle }
                    ),

                    // literal text "Table "
                    new Run(new Text("Table ") { Space = SpaceProcessingModeValues.Preserve }), // normally done by Word

                    // --- Bookmark around the SEQ field only ---
                    new BookmarkStart() { Name = substTablRef.Value, Id = "0" },
                    
                    new Run(
                        new FieldChar() { FieldCharType = FieldCharValues.Begin }),
                    new Run(
                        new FieldCode(" SEQ Table \\* ARABIC "),
                        new RunProperties(new NoProof())),
                    new Run(
                        new FieldChar() { FieldCharType = FieldCharValues.Separate }),
                    new Run(new Text("1")), // placeholder; updated by Word
                    new Run(
                        new FieldChar() { FieldCharType = FieldCharValues.End }),

                    // --- end of bookmark ---
                    new BookmarkEnd() { Id = "0" },

                    // separator and caption text
                    new Run(new Text($" – Interface operation {op?.OperationId}") { 
                        Space = SpaceProcessingModeValues.Preserve 
                    })
                    
                );

                if (_config.TableCaptionStyle != null)
                {
                    caption.ParagraphProperties = new ParagraphProperties();
                    caption.ParagraphProperties.ParagraphStyleId = new ParagraphStyleId() { Val = _config.TableCaptionStyle };
                }

                // Append to the body
                body.Append(caption);
            }

            // Really appending the table
            body.Append(table);

            // Notes
            if (opConfig?.Notes != null)
            {
                // add notes per default at the end of the table
                foreach (var note in opConfig.Notes)
                    body.AppendChild(CreateParagraph(
                        $"NOTE   {note}",
                        styleId: $"{_config.NoteStyle}"));
            }

            // empty rows
            for (int i = 0; i < _config.NumberEmptyLines; i++)
                body.AppendChild(CreateParagraph(""));
        }

        /// <summary>
        /// Export a single operation to the Word
        /// </summary>
        public void ExportOverviewOperation(
            MainDocumentPart mainPart,
            List<ExportIecInterfaceOperation.OperationTuple> opTuples)
        {
            // access
            Body? body = mainPart.Document.Body;
            if (body == null)
                return;

            // generate a table-reference
            var substTablRef = new Substitution("table-ref", $"Table{_tableRefIdCount++}", isBookmark: true);

            // Create the table
            Table table = new Table();

            // Define table properties (1 pt border, full width)
            TableProperties tblProps = new TableProperties(
                new TableWidth { Type = TableWidthUnitValues.Pct, Width = "5000" }, // 100% width (5000 = 100% in OpenXML)
                new TableLayout { Type = TableLayoutValues.Fixed }, // <=== FIXED LAYOUT
                new TableBorders(
                    new TopBorder { Val = BorderValues.Single, Size = 8 },
                    new BottomBorder { Val = BorderValues.Single, Size = 8 },
                    new LeftBorder { Val = BorderValues.Single, Size = 8 },
                    new RightBorder { Val = BorderValues.Single, Size = 8 },
                    new InsideHorizontalBorder { Val = BorderValues.Single, Size = 8 },
                    new InsideVerticalBorder { Val = BorderValues.Single, Size = 8 }
                )
            );
            table.AppendChild(tblProps);

            // Define column widths (sum ~9360 twips = ~6.5 inches)
            double cm = 567;
            int[] cw = { (int)(3 * cm), (int)(6 * cm) };

            if (_config.OverviewColumnWidthCm != null && _config.OverviewColumnWidthCm.Count >= 2)
                for (int i = 0; i < Math.Min(2, _config.OverviewColumnWidthCm.Count); i++)
                    cw[i] = (int)(cm * _config.OverviewColumnWidthCm[i]);

            TableGrid tableGrid = new TableGrid();
            foreach (int width in cw)
            {
                tableGrid.Append(new GridColumn() { Width = width.ToString() });
            }
            table.Append(tableGrid);

            // 1st row: Header 
            {
                TableRow tr = new TableRow();
                tr.Append(CreateCell("Interface operation name", cw[0], bold: true));
                tr.Append(CreateCell("Description", cw[1], bold: true));
                table.Append(tr);
            }

            // 2nd.. row: Data
            foreach (var opT in opTuples)
            {
                var explanation = opT.Config.Explanation ?? opT.Operation.Summary;
                TableRow tr = new TableRow();
                tr.Append(CreateCell($"{opT.Operation.OperationId}", cw[0]));
                tr.Append(CreateCell($"{explanation}", cw[1]));
                table.Append(tr);
            }

            // Before appending the table, add some caption text?
            if (_config.AddTableCaptions)
            {
                // Caption paragraph
                Paragraph caption = new Paragraph(
                    new ParagraphProperties(
                        new ParagraphStyleId { Val = _config.TableCaptionStyle }
                    ),

                    // literal text "Table "
                    new Run(new Text("Table ") { Space = SpaceProcessingModeValues.Preserve }), // normally done by Word

                    // --- Bookmark around the SEQ field only ---
                    new BookmarkStart() { Name = substTablRef.Value, Id = "0" },

                    new Run(
                        new FieldChar() { FieldCharType = FieldCharValues.Begin }),
                    new Run(
                        new FieldCode(" SEQ Table \\* ARABIC "),
                        new RunProperties(new NoProof())),
                    new Run(
                        new FieldChar() { FieldCharType = FieldCharValues.Separate }),
                    new Run(new Text("1")), // placeholder; updated by Word
                    new Run(
                        new FieldChar() { FieldCharType = FieldCharValues.End }),

                    // --- end of bookmark ---
                    new BookmarkEnd() { Id = "0" },

                    // separator and caption text
                    new Run(new Text($" – Overview on interface operations")
                    {
                        Space = SpaceProcessingModeValues.Preserve
                    })

                );

                if (_config.TableCaptionStyle != null)
                {
                    caption.ParagraphProperties = new ParagraphProperties();
                    caption.ParagraphProperties.ParagraphStyleId = new ParagraphStyleId() { Val = _config.TableCaptionStyle };
                }

                // Append to the body
                body.Append(caption);
            }

            // Really appending the table
            body.Append(table);

            // empty rows
            for (int i = 0; i < _config.NumberEmptyLines; i++)
                body.AppendChild(CreateParagraph(""));
        }

        //public class EmptyCollectionOmittingConverter : IYamlTypeConverter
        //{
        //    public bool Accepts(Type type)
        //    {
        //        // Handle all IEnumerable types except string
        //        return typeof(IEnumerable).IsAssignableFrom(type)
        //               && type != typeof(string);
        //    }

        //    public object? ReadYaml(IParser parser, Type type, ObjectDeserializer rootDeserializer)
        //    {
        //        // Delegate normal deserialization
        //        return rootDeserializer(type);
        //    }

        //    public void WriteYaml(IEmitter emitter, object? value, Type type, ObjectSerializer serializer)
        //    {
        //        // If null, omit
        //        if (value == null)
        //            return;

        //        // If it's a collection and EMPTY → omit
        //        if (value is IEnumerable enumerable && !enumerable.Cast<object?>().Any())
        //            return;

        //        // Otherwise (non-empty) → serialize normally
        //        serializer(value, typeof(object));
        //    }
        //}

        public sealed class OmitEmptyVisitor : ChainedObjectGraphVisitor
        {
            public OmitEmptyVisitor(IObjectGraphVisitor<IEmitter> nextVisitor)
                : base(nextVisitor)
            {
            }

            // Only override the property-level EnterMapping for class properties
            public override bool EnterMapping(IPropertyDescriptor key, IObjectDescriptor value, IEmitter context, ObjectSerializer serializer)
            {
                if (value.Value == null)
                {
                    // Skip nulls
                    return false;
                }

                // Skip empty collections (but not strings)
                if (value.Value is IEnumerable enumerable && value.Type != typeof(string))
                {
                    if (!enumerable.Cast<object?>().Any())
                    {
                        return false;
                    }
                }

                // Otherwise serialize normally
                return base.EnterMapping(key, value, context, serializer);
            }
        }

        /// <summary>
        /// Export a single operation to the Word
        /// </summary>
        public void ExportSingleYamlCode(
            MainDocumentPart mainPart,
            YamlConfig.OperationConfig opConfig,
            YamlOpenApi.OpenApiOperation origOp)
        {
            // access
            Body? body = mainPart.Document.Body;
            if (body == null)
                return;

            if (origOp == null)
                return;

            // work on a copy
            var op = origOp.Copy();

            // generate a table-reference
            var substTablRef = new Substitution("table-ref", $"Table{_tableRefIdCount++}", isBookmark: true);
            var substs = new List<Substitution>() { substTablRef };

            // try to suppress input parameters in OpenApiOperation
            if (true && op?.Parameters != null)
            {
                var sup = new List<string>();
                if (opConfig.SuppressInputs != null)
                    sup.AddRange(opConfig.SuppressInputs);
                else
                if (_config.SuppressInputs != null)
                    sup.AddRange(_config.SuppressInputs);
                
                var toDel = new List<YamlOpenApi.OpenApiParameter>();
                foreach (var si in sup) 
                    foreach (var x in op.Parameters)
                        if (si != null && true == x.Name?.Equals(si, StringComparison.InvariantCultureIgnoreCase))
                            toDel.Add(x);

                foreach (var td in toDel)
                    op.Parameters.Remove(td);
            }

            // try remove x-semanticId
            if (true && op?.SemanticIds != null)
            {
                op.SemanticIds = null;
            }

            // serialize YAML
            var serializer = new SerializerBuilder()
                .WithNamingConvention(CamelCaseNamingConvention.Instance)
                // .WithTypeConverter(new EmptyCollectionOmittingConverter())
                .ConfigureDefaultValuesHandling(DefaultValuesHandling.OmitDefaults) // omit default/null
                .WithEmissionPhaseObjectGraphVisitor(args => new OmitEmptyVisitor(args.InnerVisitor))
                .Build();
            var yaml = serializer.Serialize(op);

            // build lines
            var lines = yaml.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

            // Heading
            body.AppendChild(CreateParagraph(
                $"{opConfig?.Heading ?? _config.TableHeadingPrefix} {op.OperationId}",
                styleId: $"{_config.YamlHeadingStyle}"));

            // Intro text
            body.AppendChild(CreateParagraph(
                $"{opConfig?.Body ?? _config.Body}",
                styleId: $"{_config.BodyStyle}",
                substitutions: substs));

            // paragraphs
            var paras = CreateMonospacedParagraph(lines.ToList(), styleId: _config.YamlCodeStyle, isBoxed: true);            

            // Before appending the table, add some caption text?
            if (_config.AddTableCaptions)
            {
                // Caption paragraph
                Paragraph caption = new Paragraph(
                    new ParagraphProperties(
                        new ParagraphStyleId { Val = _config.TableCaptionStyle }
                    ),

                    // literal text "Table "
                    new Run(new Text("Table ") { Space = SpaceProcessingModeValues.Preserve }), // normally done by Word

                    // --- Bookmark around the SEQ field only ---
                    new BookmarkStart() { Name = substTablRef.Value, Id = "0" },

                    new Run(
                        new FieldChar() { FieldCharType = FieldCharValues.Begin }),
                    new Run(
                        new FieldCode(" SEQ Table \\* ARABIC "),
                        new RunProperties(new NoProof())),
                    new Run(
                        new FieldChar() { FieldCharType = FieldCharValues.Separate }),
                    new Run(new Text("1")), // placeholder; updated by Word
                    new Run(
                        new FieldChar() { FieldCharType = FieldCharValues.End }),

                    // --- end of bookmark ---
                    new BookmarkEnd() { Id = "0" },

                    // separator and caption text
                    new Run(new Text($" – Overview on interface operations")
                    {
                        Space = SpaceProcessingModeValues.Preserve
                    })

                );

                if (_config.TableCaptionStyle != null)
                {
                    caption.ParagraphProperties = new ParagraphProperties();
                    caption.ParagraphProperties.ParagraphStyleId = new ParagraphStyleId() { Val = _config.TableCaptionStyle };
                }

                // Append to the body
                body.Append(caption);
            }

            // now, append the monospaced paragraphs?
            body.Append(paras);

            // empty rows
            for (int i = 0; i < _config.NumberEmptyLines; i++)
                body.AppendChild(CreateParagraph(""));

        }

        public List<string> PatternStorage = new();

        /// <summary>
        /// Export a single schema type information with originated property bundles 
        /// </summary>
        public void ExportSinglePropertyBundle(
            YamlOpenApi.OpenApiDocument doc,
            MainDocumentPart mainPart,
            string schemaName,
            OpenApiOriginatedPropertyList oplist,
            List<string>? suppressMembers = null,
            List<string>? schemaNotFollow = null)
        {
            // access
            Body? body = mainPart.Document.Body;
            if (body == null)
                return;

            // generate a table-reference
            var substTablRef = new Substitution("table-ref", $"Table{_tableRefIdCount++}", isBookmark: true);
            var substs = (new[] { 
                substTablRef, 
                new Substitution("schema", schemaName, false) } 
            ).ToList();

            // Heading
            body.AppendChild(CreateParagraph(
                $"{_config.SchemaHeadingPrefix} {schemaName}",
                styleId: $"{_config.TableHeadingStyle}"));

            // Intro text
            body.AppendChild(CreateParagraph(
                $"{_config.SchemaBody}",
                styleId: $"{_config.BodyStyle}",
                substitutions: substs));

            // filter suppresed out
            oplist = new OpenApiOriginatedPropertyList(oplist.Where(op => !YamlOpenApi.IsContained(_config.SuppressSchemaNames, op.Name)));

            // sort according specific order
            Func<string?, int> originOrder = (str) =>
            {
                // schema itself doesget sorted very much to the bottom
                if (str?.Equals(schemaName, StringComparison.InvariantCultureIgnoreCase) == true)
                    return 99999;
                
                // if unknown, then near to bottom
                var res = 88888;

                // try to assign position
                if (str != null && _config.OriginSchemaOrder != null)
                    for (int i = 0; i < _config.OriginSchemaOrder.Count; i++)
                    {
                        if (str.Equals(_config.OriginSchemaOrder[i], StringComparison.InvariantCultureIgnoreCase))
                        {
                            res = i;
                            break;
                        }
                }

                // ok
                return res;
            };
            oplist.Sort((o1, o2) => {
                var oo1 = originOrder(YamlOpenApi.StripSchemaHead(o1.Origin));
                var oo2 = originOrder(YamlOpenApi.StripSchemaHead(o2.Origin));
                if (oo1 < oo2)
                    return -1;
                else if (oo1 > oo2)
                    return 1;
                else
                    return string.Compare(o1.Name, o2.Name, StringComparison.InvariantCultureIgnoreCase);
            });

            if (schemaName?.Contains("logicalExpression") == true)
                ;

            // for the special case oneOf .. required : check for "oneOfGroups"
            // search again for the schema
            var oneOfGroups = new Dictionary<string, int>();
            var schema = doc.FindComponent<OpenApiSchema>($"#/components/schemas/{schemaName}");
            if (schema?.OneOf != null && schema.OneOf.Count > 0)
                foreach (var oo in schema.OneOf)
                    if (oo.Required != null && oo.Required.Count > 0)
                        foreach (var member in oo.Required)
                            if (!oneOfGroups.ContainsKey(member))
                                oneOfGroups.Add(member, 1);

            // Create the table
            Table table = new Table();

            // Define table properties (1 pt border, full width)
            TableProperties tblProps = new TableProperties(
                new TableWidth { Type = TableWidthUnitValues.Pct, Width = "5000" }, // 100% width (5000 = 100% in OpenXML)
                new TableLayout { Type = TableLayoutValues.Fixed }, // <=== FIXED LAYOUT
                new TableBorders(
                    new TopBorder { Val = BorderValues.Single, Size = 8 },
                    new BottomBorder { Val = BorderValues.Single, Size = 8 },
                    new LeftBorder { Val = BorderValues.Single, Size = 8 },
                    new RightBorder { Val = BorderValues.Single, Size = 8 },
                    new InsideHorizontalBorder { Val = BorderValues.None, Size = 8 },
                    new InsideVerticalBorder { Val = BorderValues.None, Size = 8 }
                )
            );
            table.AppendChild(tblProps);

            // Define column widths (sum ~9360 twips = ~6.5 inches)
            double cm = 567;
            int[] cw = { (int)(4 * cm), (int)(4 * cm), (int)(1 * cm), (int)(1 * cm), (int)(1.5 * cm), (int)(3 * cm) };

            //if (_config.OverviewColumnWidthCm != null && _config.OverviewColumnWidthCm.Count >= 2)
            //    for (int i = 0; i < Math.Min(2, _config.OverviewColumnWidthCm.Count); i++)
            //        cw[i] = (int)(cm * _config.OverviewColumnWidthCm[i]);

            TableGrid tableGrid = new TableGrid();
            foreach (int width in cw)
            {
                tableGrid.Append(new GridColumn() { Width = width.ToString() });
            }
            table.Append(tableGrid);

            // 1st row: Top 
            if (false)
            {
                TableRow tr = new TableRow();
                tr.Append(CreateMergedCell(schemaName, true, cw[0], bold: true));
                tr.Append(CreateMergedCell("", false, cw[1]));
                tr.Append(CreateMergedCell("", false, cw[2]));
                tr.Append(CreateMergedCell("", false, cw[3]));
                tr.Append(CreateMergedCell("", false, cw[4]));
                tr.Append(CreateMergedCell("", false, cw[5]));
                table.Append(tr);
            }

            // 1st row: Column headerHeader 
            {
                TableRow tr = new TableRow();
                tr.Append(CreateCell("Member", cw[0], bold: true, verticalMerge: true, verticalMergeRestart: true));
                tr.Append(CreateCell("Choice(s) of data type", cw[1], bold: true));
                tr.Append(CreateCell("Only one", cw[2], bold: true));
                tr.Append(CreateCell("Req.", cw[3], bold: true));
                tr.Append(CreateCell("Card. if present", cw[4], bold: true));
                tr.Append(CreateCell("From", cw[5], bold: true, verticalMerge: true, verticalMergeRestart: true));
                table.Append(tr);
            }

            // 2nd row: Column headerHeader 
            {
                TableRow tr = new TableRow();
                tr.Append(CreateCell("", cw[0], bold: true, verticalMerge: true, verticalMergeRestart: false));
                tr.Append(CreateMergedCell("Further specification", true, cw[1], bold: true));
                tr.Append(CreateMergedCell("", false, cw[2], bold: true));
                tr.Append(CreateMergedCell("", false, cw[3], bold: true));
                tr.Append(CreateMergedCell("", false, cw[4], bold: true));
                tr.Append(CreateCell("", cw[4], bold: true, verticalMerge: true, verticalMergeRestart: false));
                table.Append(tr);
            }

            // 3rd.. row: single member
            string? lastFrom = null;
            foreach (var op in oplist)
            {
                //
                // Row 1 .. the usual stuff
                //

                var hasEnum = (op.Property?.Enum != null && op.Property.Enum.Count > 0);
                var hasPattern = (op.Property?.Pattern != null && op.Property.Pattern.Length > 0);
                var hasLen = (op.Property?.minLength != null || op.Property?.maxLength != null);
                var hasFormat = (op.Property?.Format != null && op.Property.Format.Length > 0);
                var needOf2ndRow = hasEnum || hasPattern || hasLen || hasFormat;

                // skip member
                if (op.Name != null && suppressMembers?.Contains(op.Name) == true)
                    continue;

                // prepare cell data, first
                var name = op.Name;
                var type = op.Property?.Type;
                if (type == null && op.Property?.Ref != null)
                    type = YamlOpenApi.StripSchemaHead(op.Property.Ref.Replace("#/components/schemas/", ""));

                var oneOf = (name != null && (oneOfGroups?.ContainsKey(name) == true) && (oneOfGroups[name] >= 0)) ? "X" : "";
                var req = "no";
                var card = "0..1";
                if (op.Required)
                {
                    req = "yes";
                    card = "1";
                }
                var from = YamlOpenApi.StripSchemaHead(op.Origin);
                if (op.Property?.Type == "array" && op.Property.Items?.Ref != null)
                {
                    type = YamlOpenApi.StripSchemaHead(op.Property.Items.Ref);
                    var min = "0";
                    var max = "*";
                    if (op.Property.MinItems != null)
                        min = op.Property.MinItems.ToString() ?? "0";
                    if (op.Property.MaxItems != null)
                        max = op.Property.MaxItems.ToString() ?? "0";
                    card = $"{min}..{max}";
                }

                // expand type??
                if (type != null && (schemaNotFollow == null || !schemaNotFollow.Contains(type)))
                {
                    var typeComp = doc.FindComponent<YamlOpenApi.OpenApiSchema>("#/components/schemas/" + type);
                    // is a one of
                    if (typeComp != null && typeComp.OneOf != null && typeComp.OneOf.Count > 0)
                    {
                        var types = new List<string>();
                        foreach (var one in typeComp.OneOf)
                            if (one?.Ref != null)
                                types.Add(YamlOpenApi.StripSchemaHead(one.Ref) ?? "?");
                        type = string.Join("\n", types);
                    }
                }

                // skip the "top" deviding line to above "from"?
                var skipDividingLine = false;
                if (lastFrom != null && lastFrom == from)
                {
                    skipDividingLine = true;
                }

                // put it in Word
                if (true) 
                {
                    TableRow tr = new TableRow();
                    tr.Append(CreateCell($"{name}", cw[0], bold: true, verticalMerge: needOf2ndRow, verticalMergeRestart: needOf2ndRow));
                    tr.Append(CreateCell($"{type}", cw[1]));
                    tr.Append(CreateCell($"{oneOf}", cw[2]));
                    tr.Append(CreateCell($"{req}", cw[3]));
                    tr.Append(CreateCell($"{card}", cw[4]));
                    tr.Append(CreateCell($"{(skipDividingLine ? "" : from)}", cw[5], verticalMerge: true, verticalMergeRestart: !skipDividingLine));
                    table.Append(tr);

                    // if skipped dividing line, do not break across pages here
                    if (skipDividingLine)
                    {
                        TableRowProperties trPr = new TableRowProperties(
                            new CantSplit() // prevent row from splitting across pages
                        );
                        tr.Append(trPr);
                    }
                }

                // state
                lastFrom = from;

                //
                // Row 2 : extras .. enum?
                // Challenge: keep the 5th column, as it is ..
                //

                if (needOf2ndRow)
                {
                    // basically evaluate, what is in the 2nd row
                    var secondText = "";
                    
                    if (hasEnum && op.Property?.Enum != null)
                    {
                        var types = new List<string>();
                        types.Add("Enumeration values: ");
                        foreach (var etxt in op.Property.Enum)
                            if (etxt != null)
                                types.Add($"\u2014 {etxt}");
                        secondText = string.Join("\n", types);
                    }

                    if (hasPattern && op.Property?.Pattern != null)
                    {
                        // depict pattern directly or store for later?
                        if (op.Property.Pattern.Length <= _config.PatternInlineLimit)
                        {
                            // direct
                            if (secondText.Length > 0)
                                secondText += "\n";
                            secondText += $"Pattern: {op.Property.Pattern}";
                        }
                        else
                        {
                            // already contained? .. or add?
                            var patternNdx = PatternStorage.IndexOf(op.Property.Pattern.Trim());
                            if (patternNdx < 0)
                            {
                                patternNdx = PatternStorage.Count;
                                PatternStorage.Add(op.Property.Pattern);
                            }

                            // write out Index
                            if (secondText.Length > 0)
                                secondText += "\n";
                            secondText += $"See string pattern index {1 + patternNdx}";
                        }
                    }

                    if (hasLen)
                    {
                        if (secondText.Length > 0)
                            secondText += "\n";
                        if (op.Property?.minLength != null)
                            secondText += $"Minimum length: {op.Property.minLength.ToString()} ";
                        if (op.Property?.maxLength != null)
                            secondText += $"Maximum length: {op.Property.maxLength.ToString()} ";
                    }

                    if (hasFormat)
                    {
                        if (secondText.Length > 0)
                            secondText += "\n";
                        secondText += $"Format: {op.Property?.Format} ";
                    }

                    // put it in Word (basically it is only about <secondText>)
                    TableRow tr = new TableRow();
                    tr.Append(CreateCell($"", cw[0], verticalMerge: true, verticalMergeRestart: false));
                    tr.Append(CreateMergedCell(secondText, true, cw[1]));
                    tr.Append(CreateMergedCell("", false, cw[2]));
                    tr.Append(CreateMergedCell("", false, cw[3]));
                    tr.Append(CreateMergedCell("", false, cw[4]));
                    tr.Append(CreateCell($"", cw[5], verticalMerge: true, verticalMergeRestart: false));
                    // tr.Append(CreateMergedCell("", false, cw[4]));
                    table.Append(tr);

                    // if skipped dividing line, do not break across pages here
                    if (skipDividingLine)
                    {
                        TableRowProperties trPr = new TableRowProperties(
                            new CantSplit() // prevent row from splitting across pages
                        );
                        tr.Append(trPr);
                    }
                }
            }

            // last row: Key
            {
                TableRow tr = new TableRow();
                tr.Append(CreateMergedCell("Key of table:\nReq.:    Required.\nCard. if present:    Minimum and maximum cardinality of the JSON value, if value is present.", true, cw[0]));
                tr.Append(CreateMergedCell("", false, cw[1]));
                tr.Append(CreateMergedCell("", false, cw[2]));
                tr.Append(CreateMergedCell("", false, cw[3]));
                tr.Append(CreateMergedCell("", false, cw[4]));
                tr.Append(CreateMergedCell("", false, cw[5]));
                table.Append(tr);
            }

            // Before appending the table, add some caption text?
            if (_config.AddTableCaptions)
            {
                // Caption paragraph
                Paragraph caption = new Paragraph(
                    new ParagraphProperties(
                        new ParagraphStyleId { Val = _config.TableCaptionStyle }
                    ),

                    // literal text "Table "
                    new Run(new Text("Table ") { Space = SpaceProcessingModeValues.Preserve }), // normally done by Word

                    // --- Bookmark around the SEQ field only ---
                    new BookmarkStart() { Name = substTablRef.Value, Id = "0" },

                    new Run(
                        new FieldChar() { FieldCharType = FieldCharValues.Begin }),
                    new Run(
                        new FieldCode(" SEQ Table \\* ARABIC "),
                        new RunProperties(new NoProof())),
                    new Run(
                        new FieldChar() { FieldCharType = FieldCharValues.Separate }),
                    new Run(new Text("1")), // placeholder; updated by Word
                    new Run(
                        new FieldChar() { FieldCharType = FieldCharValues.End }),

                    // --- end of bookmark ---
                    new BookmarkEnd() { Id = "0" },

                    // separator and caption text
                    new Run(new Text($" – {_config.SchemaTableCaptionPrefix} {schemaName}")
                    {
                        Space = SpaceProcessingModeValues.Preserve
                    })

                );

                if (_config.TableCaptionStyle != null)
                {
                    caption.ParagraphProperties = new ParagraphProperties();
                    caption.ParagraphProperties.ParagraphStyleId = new ParagraphStyleId() { Val = _config.TableCaptionStyle };
                }

                // Append to the body
                body.Append(caption);
            }

            // Really appending the table
            body.Append(table);

            // empty rows
            for (int i = 0; i < _config.NumberEmptyLines; i++)
                body.AppendChild(CreateParagraph(""));
        }

        /// <summary>
        /// Export a table with the stored (long) patterns
        /// </summary>
        public void ExportPatternStorage(
            YamlOpenApi.OpenApiDocument doc,
            MainDocumentPart mainPart)
        {
            // access
            Body? body = mainPart.Document.Body;
            if (body == null)
                return;

            // generate a table-reference
            var substTablRef = new Substitution("table-ref", $"Table{_tableRefIdCount++}", isBookmark: true);
            var substs = (new[] { substTablRef } ).ToList();

            // Create the table
            Table table = new Table();

            // Define table properties (1 pt border, full width)
            TableProperties tblProps = new TableProperties(
                new TableWidth { Type = TableWidthUnitValues.Pct, Width = "5000" }, // 100% width (5000 = 100% in OpenXML)
                new TableLayout { Type = TableLayoutValues.Fixed }, // <=== FIXED LAYOUT
                new TableBorders(
                    new TopBorder { Val = BorderValues.Single, Size = 8 },
                    new BottomBorder { Val = BorderValues.Single, Size = 8 },
                    new LeftBorder { Val = BorderValues.Single, Size = 8 },
                    new RightBorder { Val = BorderValues.Single, Size = 8 },
                    new InsideHorizontalBorder { Val = BorderValues.None, Size = 8 },
                    new InsideVerticalBorder { Val = BorderValues.None, Size = 8 }
                )
            );
            table.AppendChild(tblProps);

            // Define column widths (sum ~9360 twips = ~6.5 inches)
            double cm = 567;
            int[] cw = { (int)(2 * cm), (int)(17 * cm) };

            //if (_config.OverviewColumnWidthCm != null && _config.OverviewColumnWidthCm.Count >= 2)
            //    for (int i = 0; i < Math.Min(2, _config.OverviewColumnWidthCm.Count); i++)
            //        cw[i] = (int)(cm * _config.OverviewColumnWidthCm[i]);

            TableGrid tableGrid = new TableGrid();
            foreach (int width in cw)
            {
                tableGrid.Append(new GridColumn() { Width = width.ToString() });
            }
            table.Append(tableGrid);

            // 1st row: Column headerHeader 
            {
                TableRow tr = new TableRow();
                tr.Append(CreateCell("Index", cw[0]));
                tr.Append(CreateCell("String pattern definition", cw[1]));
                table.Append(tr);
            }

            // 2nd.. row: single member
            for (int pi = 0; pi < PatternStorage.Count; pi++)
            {
                var pat = PatternStorage[pi];
                
                // put it in Word
                TableRow tr = new TableRow();
                tr.Append(CreateCell($"{1+pi}", cw[0], bold: true));
                tr.Append(CreateCell($"{pat}", cw[1]));
                table.Append(tr);
            }

            // Before appending the table, add some caption text?
            if (_config.AddTableCaptions)
            {
                // Caption paragraph
                Paragraph caption = new Paragraph(
                    new ParagraphProperties(
                        new ParagraphStyleId { Val = _config.TableCaptionStyle }
                    ),

                    // literal text "Table "
                    new Run(new Text("Table ") { Space = SpaceProcessingModeValues.Preserve }), // normally done by Word

                    // --- Bookmark around the SEQ field only ---
                    new BookmarkStart() { Name = substTablRef.Value, Id = "0" },

                    new Run(
                        new FieldChar() { FieldCharType = FieldCharValues.Begin }),
                    new Run(
                        new FieldCode(" SEQ Table \\* ARABIC "),
                        new RunProperties(new NoProof())),
                    new Run(
                        new FieldChar() { FieldCharType = FieldCharValues.Separate }),
                    new Run(new Text("1")), // placeholder; updated by Word
                    new Run(
                        new FieldChar() { FieldCharType = FieldCharValues.End }),

                    // --- end of bookmark ---
                    new BookmarkEnd() { Id = "0" },

                    // separator and caption text
                    new Run(new Text($" – {_config.PatternTableCaptionPrefix}")
                    {
                        Space = SpaceProcessingModeValues.Preserve
                    })

                );

                if (_config.TableCaptionStyle != null)
                {
                    caption.ParagraphProperties = new ParagraphProperties();
                    caption.ParagraphProperties.ParagraphStyleId = new ParagraphStyleId() { Val = _config.TableCaptionStyle };
                }

                // Append to the body
                body.Append(caption);
            }

            // Really appending the table
            body.Append(table);

            // clear storage
            PatternStorage.Clear();

            // empty rows
            for (int i = 0; i < _config.NumberEmptyLines; i++)
                body.AppendChild(CreateParagraph(""));
        }

        /// <summary>
        /// Export a single schema type information with originated property bundles 
        /// </summary>
        public void ExportSingleHttpOperationDescription(
            YamlOpenApi.OpenApiDocument doc,
            MainDocumentPart mainPart,
            YamlConfig.OperationConfig opConfig,
            YamlOpenApi.OpenApiOperation origOp)
        {
            // access
            Body? body = mainPart.Document.Body;
            if (body == null)
                return;

            if (origOp == null)
                return;

            // work on a copy
            var op = origOp.Copy();

            // try to suppress input parameters in OpenApiOperation (have exclusive)
            if (true && op?.Parameters != null)
            {
                var sup = new List<string>();
                if (opConfig.SuppressInputs != null)
                    sup.AddRange(opConfig.SuppressInputs);
                else 
                if (_config.SuppressInputs != null)
                    sup.AddRange(_config.SuppressInputs);

                var toDel = new List<YamlOpenApi.OpenApiParameter>();
                foreach (var si in sup)
                    foreach (var x in op.Parameters)
                        if (si != null && true == x.Name?.Equals(si, StringComparison.InvariantCultureIgnoreCase))
                            toDel.Add(x);

                foreach (var td in toDel)
                    op.Parameters.Remove(td);
            }

            // try remove x-semanticId
            if (true && op?.SemanticIds != null)
            {
                op.SemanticIds = null;
            }

            // build explanation
            var explanation = opConfig?.Explanation ?? op?.Summary;

            // need a lambda for filling up responses, later
            Func<OpenApiResponse, Tuple<OpenApiResponse, string?>> lambdaCreateDereferencedWorkResponse = (resp) =>
            {
                var res = resp.Clone();
                string? pureSchemaName = null;

                // try fill some layout decision
                if (resp.Ref != null)
                {
                    var r = doc.FindComponent<YamlOpenApi.OpenApiResponse>(resp.Ref);
                    if (r != null)
                        res.Join(r);
                    else
                        pureSchemaName = YamlOpenApi.StripResponseHead(resp.Ref);
                }

                return new Tuple<OpenApiResponse, string?>(res, pureSchemaName);
            };

            // generate a table-reference
            var substTablRef = new Substitution("table-ref", $"Table{_tableRefIdCount++}", isBookmark: true);
            var substs = (new[] {
                substTablRef,
                new Substitution("operation-id", op?.OperationId ?? "", false) }
            ).ToList();

            // Heading
            body.AppendChild(CreateParagraph(
                $"{opConfig?.Heading ?? _config.TableHeadingPrefix} {op?.OperationId}",
                styleId: $"{_config.YamlHeadingStyle}"));

            // Intro text
            body.AppendChild(CreateParagraph(
                $"{opConfig?.Body ?? _config.Body}",
                styleId: $"{_config.BodyStyle}",
                substitutions: substs));

            // Create the table
            Table table = new Table();

            // Define table properties (1 pt border, full width)
            TableProperties tblProps = new TableProperties(
                new TableWidth { Type = TableWidthUnitValues.Pct, Width = "5000" }, // 100% width (5000 = 100% in OpenXML)
                new TableLayout { Type = TableLayoutValues.Fixed }, // <=== FIXED LAYOUT
                new TableBorders(
                    new TopBorder { Val = BorderValues.Single, Size = 8 },
                    new BottomBorder { Val = BorderValues.Single, Size = 8 },
                    new LeftBorder { Val = BorderValues.Single, Size = 8 },
                    new RightBorder { Val = BorderValues.Single, Size = 8 },
                    new InsideHorizontalBorder { Val = BorderValues.None, Size = 8 },
                    new InsideVerticalBorder { Val = BorderValues.None, Size = 8 }
                )
            );
            table.AppendChild(tblProps);

            // Approach: evaluate, if a 3- or 5-column design is required
            var fiveCols = false;
            if (op?.Responses != null)
                foreach (var resp in op.Responses)
                    if (resp.Value != null)
                    {
                        var workResp = lambdaCreateDereferencedWorkResponse(resp.Value);
                        if (workResp.Item1.Headers != null && workResp.Item1.Headers.Count > 0)
                            fiveCols = true;
                    }
            

            // Define column widths (sum ~9360 twips = ~6.5 inches)
            double cm = 567;
            var cw = new List<int>(new []{ (int)(3 * cm), (int)(3 * cm), (int)(3 * cm), (int)(3 * cm), (int)(8 * cm) });

            if (_config.InterfaceOpFiveColumnWidthCm != null && _config.InterfaceOpFiveColumnWidthCm.Count >= 5)
                for (int i = 0; i < Math.Min(5, _config.InterfaceOpFiveColumnWidthCm.Count); i++)
                    cw[i] = (int)(cm * _config.InterfaceOpFiveColumnWidthCm[i]);

            if (!fiveCols)
            {
                cw = new List<int>(new[] { (int)(3 * cm), (int)(3 * cm), (int)(14 * cm), (int)(0.1 * cm), (int)(0.1 * cm) });

                if (_config.InterfaceOpThreeColumnWidthCm != null && _config.InterfaceOpThreeColumnWidthCm.Count >= 5)
                    for (int i = 0; i < Math.Min(5, _config.InterfaceOpThreeColumnWidthCm.Count); i++)
                        cw[i] = (int)(cm * _config.InterfaceOpThreeColumnWidthCm[i]);
            }

            TableGrid tableGrid = new TableGrid();
            foreach (int width in cw)
            {
                tableGrid.Append(new GridColumn() { Width = width.ToString() });
            }
            table.Append(tableGrid);

            //
            // Start
            //

            // 1st row: Header for interface operation
            {
                TableRow tr = new TableRow();
                tr.Append(CreateCell("Interface Operation Name ", cw[0]));
                tr.Append(CreateMergedCell($"{op?.OperationId}", true, cw[1], bold: true));
                tr.Append(CreateMergedCell("", false, cw[2]));
                tr.Append(CreateMergedCell("", false, cw[3]));
                tr.Append(CreateMergedCell("", false, cw[4]));
                table.Append(tr);
            }

            // 2nd row: Explanation
            {
                TableRow tr = new TableRow();
                tr.Append(CreateCell("Explanation", cw[0]));
                tr.Append(CreateMergedCell($"{explanation}", true, cw[1], bold: true));
                tr.Append(CreateMergedCell("", false, cw[2]));
                tr.Append(CreateMergedCell("", false, cw[3]));
                tr.Append(CreateMergedCell("", false, cw[4]));
                table.Append(tr);
            }

            //
            // Request parameters
            //

            if (op?.Parameters != null && op.Parameters.Count > 0)
            {
                // Header for request parameters
                {
                    TableRow tr = new TableRow();
                    tr.Append(CreateMergedCell($"Request parameters", true, cw[0], bold: true));
                    tr.Append(CreateMergedCell("", false, cw[1]));
                    tr.Append(CreateMergedCell("", false, cw[2]));
                    tr.Append(CreateMergedCell("", false, cw[3]));
                    tr.Append(CreateMergedCell("", false, cw[4]));
                    table.Append(tr);
                }

                foreach (var para in op.Parameters)
                {
                    // 1st row is always Name + Description + Description value
                    if (true)
                    {
                        TableRow tr = new TableRow();
                        tr.Append(CreateCell($"{para.Name}", cw[0], bold: true, verticalMerge: true, verticalMergeRestart: true));
                        tr.Append(CreateMergedCell("Description", false, cw[1]));
                        tr.Append(CreateMergedCell($"{para.Description}", true, cw[2]));
                        tr.Append(CreateMergedCell("", false, cw[3]));
                        tr.Append(CreateMergedCell("", false, cw[4]));
                        table.Append(tr);
                    }

                    // 2st row is always (continued) + Required
                    if (true)
                    {
                        TableRow tr = new TableRow();
                        tr.Append(CreateCell("", cw[0], verticalMerge: true, verticalMergeRestart: false));
                        tr.Append(CreateMergedCell("Required", false, cw[1]));
                        tr.Append(CreateMergedCell($"{(para.Required ? "yes" : "no")}", true, cw[2]));
                        tr.Append(CreateMergedCell("", false, cw[3]));
                        tr.Append(CreateMergedCell("", false, cw[4]));
                        table.Append(tr);
                    }
                }
            }

            //
            // Request body
            //

            if (op?.RequestBody?.Content != null && op?.RequestBody.Content.Count > 0)
            {
                // Header for request body
                {
                    TableRow tr = new TableRow();
                    tr.Append(CreateMergedCell($"Request body", true, cw[0], bold: true));
                    tr.Append(CreateMergedCell("", false, cw[1]));
                    tr.Append(CreateMergedCell("", false, cw[2]));
                    tr.Append(CreateMergedCell("", false, cw[3]));
                    tr.Append(CreateMergedCell("", false, cw[4]));
                    table.Append(tr);
                }

                // Description
                if (true)
                {
                    TableRow tr = new TableRow();
                    tr.Append(CreateCell("Description", cw[0]));
                    tr.Append(CreateMergedCell($"{op.RequestBody.Description}", true, cw[1], bold: true));
                    tr.Append(CreateMergedCell("", false, cw[2]));
                    tr.Append(CreateMergedCell("", false, cw[3]));
                    tr.Append(CreateMergedCell("", false, cw[4]));
                    table.Append(tr);
                }

                // Required
                if (true)
                {
                    TableRow tr = new TableRow();
                    tr.Append(CreateCell("Required", cw[0]));
                    tr.Append(CreateMergedCell($"{(op.RequestBody.Required ? "yes" : "no")}", true, cw[1], bold: true));
                    tr.Append(CreateMergedCell("", false, cw[2]));
                    tr.Append(CreateMergedCell("", false, cw[3]));
                    tr.Append(CreateMergedCell("", false, cw[4]));
                    table.Append(tr);
                }

                // Invisible to the reader: multiple content types
                foreach (var cntTup in op.RequestBody.Content)
                {
                    if (cntTup.Key == null || cntTup.Value == null)
                        continue;

                    // Content type
                    if (true)
                    {
                        TableRow tr = new TableRow();
                        tr.Append(CreateCell("Content type", cw[0]));
                        tr.Append(CreateMergedCell($"{cntTup.Key}", true, cw[1], bold: true));
                        tr.Append(CreateMergedCell("", false, cw[2]));
                        tr.Append(CreateMergedCell("", false, cw[3]));
                        tr.Append(CreateMergedCell("", false, cw[4]));
                        table.Append(tr);
                    }

                    // Schema
                    if (cntTup.Value.Schema?.Ref != null)
                    {
                        var schema = YamlOpenApi.StripSchemaHead(cntTup.Value.Schema.Ref);

                        TableRow tr = new TableRow();
                        tr.Append(CreateCell("Schema", cw[0]));
                        tr.Append(CreateMergedCell($"{schema}", true, cw[1], bold: true));
                        tr.Append(CreateMergedCell("", false, cw[2]));
                        tr.Append(CreateMergedCell("", false, cw[3]));
                        tr.Append(CreateMergedCell("", false, cw[4]));
                        table.Append(tr);
                    }
                }
            }

            //
            // Responses
            //

            if (op?.Responses != null && op.Responses.Count > 0)
            {
                // Header for responses
                {
                    TableRow tr = new TableRow();
                    tr.Append(CreateMergedCell($"Responses", true, cw[0], bold: true));
                    tr.Append(CreateMergedCell("", false, cw[1]));
                    tr.Append(CreateMergedCell("", false, cw[2]));
                    tr.Append(CreateMergedCell("", false, cw[3]));
                    tr.Append(CreateMergedCell("", false, cw[4]));
                    table.Append(tr);
                }

                // Column header for responses
                {
                    TableRow tr = new TableRow();
                    tr.Append(CreateCell("Status code", cw[0], bold: true));
                    tr.Append(CreateMergedCell($"Further information", true, cw[1]));
                    tr.Append(CreateMergedCell("", false, cw[2]));
                    tr.Append(CreateMergedCell("", false, cw[3]));
                    tr.Append(CreateMergedCell("", false, cw[4]));
                    table.Append(tr);
                }                

                // multiple responses
                foreach (var resp in op.Responses)
                {
                    // working response
                    if (resp.Value == null)
                        continue;
                    var workResp = lambdaCreateDereferencedWorkResponse(resp.Value);

                    // what to show?
                    var RefOrDescKey = "";
                    var RefOrDescVal = "";
                    if (workResp.Item1.Description != null)
                    {
                        RefOrDescKey = "Description";
                        RefOrDescVal = "" + workResp.Item1.Description;
                    }
                    else
                    if (workResp.Item2 != null)
                    {
                        RefOrDescKey = "Response";
                        RefOrDescVal = "" + workResp.Item2;
                    }
                    else
                    {
                        Console.WriteLine($"ERROR!! In {op.OperationId} no valid general info for response {resp.Key}");
                        continue;
                    }

                    // HTTP Status code + Description
                    {
                        TableRow tr = new TableRow();
                        tr.Append(CreateCell($"{resp.Key}", cw[0], bold: true));
                        tr.Append(CreateCell($"{RefOrDescKey}", cw[1]));
                        tr.Append(CreateMergedCell($"{RefOrDescVal}", true, cw[2], bold: true));
                        tr.Append(CreateMergedCell("", false, cw[3]));
                        tr.Append(CreateMergedCell("", false, cw[4]));
                        table.Append(tr);
                    }

                    if (workResp.Item1.Headers != null && workResp.Item1.Headers.Count > 0)
                    {
                        foreach (var header in workResp.Item1.Headers) {
                            // Header / name / Description
                            if (true)
                            {
                                TableRow tr = new TableRow();
                                tr.Append(CreateCell($"", cw[0]));
                                tr.Append(CreateCell($"Header", cw[1]));
                                tr.Append(CreateCell($"{header.Key}", cw[2], bold: true));
                                tr.Append(CreateCell($"Description", cw[3]));
                                tr.Append(CreateMergedCell($"{header.Value?.Description}", true, cw[4], bold: true));
                                table.Append(tr);
                            }

                            // Header / Schema
                            var dt = YamlOpenApi.StripSchemaHead(header.Value?.Schema?.Ref)
                                     ?? header.Value?.Schema?.Type;

                            if (dt != null)
                            {
                                TableRow tr = new TableRow();
                                tr.Append(CreateCell($"", cw[0]));
                                tr.Append(CreateCell($"", cw[1]));
                                tr.Append(CreateCell($"", cw[2]));
                                tr.Append(CreateCell($"Data type", cw[3]));
                                tr.Append(CreateMergedCell($"{dt}", true, cw[4], bold: true));
                                table.Append(tr);
                            }
                        }
                    }

                    // have a content description in this response
                    if (workResp.Item1.Content != null && workResp.Item1.Content.Count > 0)
                        foreach (var respCont in workResp.Item1.Content)
                        {
                            if (respCont.Value == null)
                                continue;

                            // write out content type
                            if (respCont.Key != null)
                            {
                                TableRow tr = new TableRow();
                                tr.Append(CreateCell("", cw[0]));
                                tr.Append(CreateCell("Content type", cw[1]));
                                tr.Append(CreateMergedCell($"{respCont.Key}", true, cw[2]));
                                tr.Append(CreateMergedCell("", false, cw[3]));
                                tr.Append(CreateMergedCell("", false, cw[4]));
                                table.Append(tr);
                            }

                            // the schema info should link to a schema, e.g. AssetAdministrationShell (200)
                            // or Result (401). For the first, it would be difficult to go further here, therefore
                            // simply print out ref nam

                            if (respCont.Value?.Schema?.Ref != null)
                            {
                                var rf = YamlOpenApi.StripSchemaHead(respCont.Value.Schema.Ref);

                                TableRow tr = new TableRow();
                                tr.Append(CreateMergedCell("", false, cw[0]));
                                tr.Append(CreateCell("Data type", cw[1]));
                                tr.Append(CreateMergedCell($"{rf}", true, cw[2]));
                                tr.Append(CreateMergedCell("", false, cw[3]));
                                tr.Append(CreateMergedCell("", false, cw[4]));
                                table.Append(tr);
                            }
                            else
                            {
                                Console.WriteLine($"ERROR!! In {op.OperationId} no valid schema info for response {resp.Key}");
                            }
                        }
                }
            }            

            //
            // Before appending the table, add some caption text?
            //
            if (_config.AddTableCaptions)
            {
                // Caption paragraph
                Paragraph caption = new Paragraph(
                    new ParagraphProperties(
                        new ParagraphStyleId { Val = _config.TableCaptionStyle }
                    ),

                    // literal text "Table "
                    new Run(new Text("Table ") { Space = SpaceProcessingModeValues.Preserve }), // normally done by Word

                    // --- Bookmark around the SEQ field only ---
                    new BookmarkStart() { Name = substTablRef.Value, Id = "0" },

                    new Run(
                        new FieldChar() { FieldCharType = FieldCharValues.Begin }),
                    new Run(
                        new FieldCode(" SEQ Table \\* ARABIC "),
                        new RunProperties(new NoProof())),
                    new Run(
                        new FieldChar() { FieldCharType = FieldCharValues.Separate }),
                    new Run(new Text("1")), // placeholder; updated by Word
                    new Run(
                        new FieldChar() { FieldCharType = FieldCharValues.End }),

                    // --- end of bookmark ---
                    new BookmarkEnd() { Id = "0" },

                    // separator and caption text
                    new Run(new Text($" – {_config.SchemaTableCaptionPrefix} {op?.OperationId}")
                    {
                        Space = SpaceProcessingModeValues.Preserve
                    })

                );

                if (_config.TableCaptionStyle != null)
                {
                    caption.ParagraphProperties = new ParagraphProperties();
                    caption.ParagraphProperties.ParagraphStyleId = new ParagraphStyleId() { Val = _config.TableCaptionStyle };
                }

                // Append to the body
                body.Append(caption);
            }

            // Really appending the table
            body.Append(table);

            // empty rows
            for (int i = 0; i < _config.NumberEmptyLines; i++)
                body.AppendChild(CreateParagraph(""));
        }

        //
        // Helpers for OpenXML / Word
        // 

        static Paragraph CreateParagraph (
            string text,
            string? styleId = null,
            List<Substitution>? substitutions = null)
        {
            var p = new Paragraph();

            if (styleId != null)
            {
                p.ParagraphProperties = new ParagraphProperties();
                p.ParagraphProperties.ParagraphStyleId = new ParagraphStyleId() { Val = styleId };
            }

            // chunk apart the text
            var rest = text;
            if (substitutions != null)
            {
                while (rest.Length > 0)
                {
                    // check if there is a placeholder
                    var found = false;
                    foreach (var sub in substitutions)
                    {
                        var i = rest.IndexOf($"%{sub.Key}%");
                        if (i >= 0)
                        {
                            // found
                            found = true;

                            // process the first part as a Run
                            var first = rest.Substring(0, i);
                            p.Append(new Run(new Text(first) { Space = SpaceProcessingModeValues.Preserve }));

                            // add the placeholder value
                            if (!sub.isBookmark)
                            {
                                // just add a Run
                                p.Append(new Run(new Text(sub.Value) { Space = SpaceProcessingModeValues.Preserve }));
                            }
                            else
                            {
                                // add a bookmark reference
                                p.Append(
                                    new Run(new FieldChar() { FieldCharType = FieldCharValues.Begin }),
                                    new Run(new FieldCode($" REF {sub.Value} \\h ")),
                                    new Run(new FieldChar() { FieldCharType = FieldCharValues.Separate }),
                                    new Run(new Text("1")),  // placeholder, updated by Word
                                    new FieldChar() { FieldCharType = FieldCharValues.End }
                                );
                            }

                            // let the rest by the rest after the placeholder
                            rest = rest.Substring(i + 2 + sub.Key.Length);

                            break;
                        }
                    }

                    // if not found, eat up the rest
                    if (!found)
                    {
                        p.Append(new Run(new Text(rest) { Space = SpaceProcessingModeValues.Preserve }));
                        break;
                    }
                }
            }
            else
            {
                // just one Run
                p.Append(new Run(new Text(text)));
            }

            return p;
        }

        public List<Paragraph> CreateMonospacedParagraph(
            List<string> lines,
            string? styleId = null,
            bool isBoxed = false)
        {
            // try choose one Paragraph with multiple Runs
            var firstPara = new Paragraph();

            // Add optional style (e.g., "Normal", "CodeBlock", etc.)
            ParagraphProperties pPr = new ParagraphProperties();
            if (!string.IsNullOrEmpty(styleId))
                pPr.Append(new ParagraphStyleId() { Val = styleId });

            // Optional: add paragraph borders if boxed
            if (isBoxed)
            {
                var bw = new DocumentFormat.OpenXml.UInt32Value(_config.YamlMonoBorderWidth);
                ParagraphBorders borders = new ParagraphBorders(
                    new TopBorder() { Val = BorderValues.Single, Size = bw },
                    new BottomBorder() { Val = BorderValues.Single, Size = bw },
                    new LeftBorder() { Val = BorderValues.Single, Size = bw },
                    new RightBorder() { Val = BorderValues.Single, Size = bw }
                );
                pPr.Append(borders);

                // Optional spacing inside (acts like padding)
                pPr.Append(new SpacingBetweenLines()
                {
                    Before = "120", // 6pt top space
                    After = "120"   // 6pt bottom space
                });
                pPr.Append(new Indentation() { Left = "0", Right = "0" });
            }

            firstPara.ParagraphProperties = pPr;

            // fill content
            for (int i=0; i<lines.Count; i++)
            {
                var line = lines[i];

                // Add font + size formatting
                Run run = new Run();
                RunProperties runProps = new RunProperties(
                    new RunFonts { Ascii = "CourierNew", HighAnsi = "CourierNew" },
                    new FontSize { Val = "16" } // 8 pt
                );

                run.Append(runProps);
                run.Append(new Text(line ?? "") { Space = SpaceProcessingModeValues.Preserve });

                if (i < lines.Count - 1)
                    run.Append(new Break());

                firstPara.Append(run);
            }

            return new List<Paragraph>(new[] { firstPara });
        }

        public TableCell CreateCell(string text, int width, bool bold = false,
            bool verticalMerge = false, bool verticalMergeRestart = true)
        {
            TableCell cell = new TableCell();

            var bw = new DocumentFormat.OpenXml.UInt32Value(_config.TableCellBorderWidth);

            TableCellProperties cellProps = new TableCellProperties(
                new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = width.ToString() },
                new TableCellBorders(
                    new TopBorder { Val = BorderValues.Single, Size = bw },
                    new BottomBorder { Val = BorderValues.Single, Size = bw },
                    new LeftBorder { Val = BorderValues.Single, Size = bw },
                    new RightBorder { Val = BorderValues.Single, Size = bw }
                ),
                new TableCellMargin(
                    new LeftMargin { Width = "80", Type = TableWidthUnitValues.Dxa },
                    new RightMargin { Width = "80", Type = TableWidthUnitValues.Dxa },
                    new TopMargin { Width = "40", Type = TableWidthUnitValues.Dxa },
                    new BottomMargin { Width = "40", Type = TableWidthUnitValues.Dxa }
                ),
                new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Top }
            );

            if (verticalMerge)
            {
                var vm = new VerticalMerge();
                vm.Val = verticalMergeRestart ? MergedCellValues.Restart : MergedCellValues.Continue;
                cellProps.AddChild(vm);
            }
            cell.Append(cellProps);

            // prepare Paragraph
            Paragraph paragraph = new Paragraph()
            {
                ParagraphProperties = new ParagraphProperties(new SpacingBetweenLines { After = "120" })
            };

            // split in multiple lines by '\n'
            var lines = (text ?? "").Split(new[] { '\n' }, StringSplitOptions.None);
            foreach (var line in lines)
            {
                // Add font + size formatting
                Run run = new Run();
                RunProperties runProps = new RunProperties(
                    new RunFonts { Ascii = "Arial", HighAnsi = "Arial" },
                    new FontSize { Val = "16" } // 8 pt
                );

                if (bold)
                    runProps.Append(new Bold());

                run.Append(runProps);
                run.Append(new Text(line) { Space = SpaceProcessingModeValues.Preserve });
                if (line != lines.Last())
                    run.Append(new Break());

                paragraph.Append(run);
            }
            cell.Append(paragraph);
                        
            return cell;
        }

        public TableCell CreateMergedCell(string text, bool isStart, int width, bool bold = false)
        {
            TableCell cell = new TableCell();

            var bw = new DocumentFormat.OpenXml.UInt32Value(_config.TableCellBorderWidth);

            TableCellProperties cellProps = new TableCellProperties(
                new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = width.ToString() },
                new HorizontalMerge { Val = isStart ? MergedCellValues.Restart : MergedCellValues.Continue },
                new TableCellBorders(
                    new TopBorder { Val = BorderValues.Single, Size = bw },
                    new BottomBorder { Val = BorderValues.Single, Size = bw },
                    new LeftBorder { Val = BorderValues.Single, Size = bw },
                    new RightBorder { Val = BorderValues.Single, Size = bw }
                ),
                new TableCellMargin(
                    new LeftMargin { Width = "80", Type = TableWidthUnitValues.Dxa },
                    new RightMargin { Width = "80", Type = TableWidthUnitValues.Dxa },
                    new TopMargin { Width = "40", Type = TableWidthUnitValues.Dxa },
                    new BottomMargin { Width = "40", Type = TableWidthUnitValues.Dxa }
                ),
                new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Top }
            );
            cell.Append(cellProps);

            // prepare Paragraph
            Paragraph paragraph = new Paragraph()
            {
                ParagraphProperties = new ParagraphProperties(new SpacingBetweenLines { After = "120" })
            };

            // split in multiple lines by '\n'
            var lines = (text ?? "").Split(new[] { '\n' }, StringSplitOptions.None);
            foreach (var line in lines)
            {
                // Add font + size formatting
                Run run = new Run();
                RunProperties runProps = new RunProperties(
                    new RunFonts { Ascii = "Arial", HighAnsi = "Arial" },
                    new FontSize { Val = "16" } // 8 pt
                );

                if (bold)
                    runProps.Append(new Bold());

                run.Append(runProps);
                run.Append(new Text(line) { Space = SpaceProcessingModeValues.Preserve });
                if (line != lines.Last())
                    run.Append(new Break());

                paragraph.Append(run);
            }
            cell.Append(paragraph);

            return cell;
        }

        public static void ListStyleNames(MainDocumentPart? mainPart, string prefix = "")
        {
            var stylesPart = mainPart?.StyleDefinitionsPart;
            if (stylesPart?.Styles == null)
            {
                Console.WriteLine($"{prefix}No styles found in the document.");
                return;
            }

            var styles = stylesPart.Styles.Elements<Style>();

            Console.WriteLine($"{prefix}Found {styles.Count()} styles.");

            foreach (var style in styles)
            {
                string styleId = style.StyleId?.ToString() ?? "(no ID)";
                string name = style.StyleName?.Val?.ToString() ?? "(no name)";
                string type = style.Type?.ToString() ?? "(no type)";

                Console.WriteLine($"{styleId,-30} | {name,-40} | {type}");
            }
        }

        public static void GenerateDefaultStyles(StyleDefinitionsPart stylePart)
        {
            Styles styles = new Styles();

            // Normal
            Style normalStyle = new Style()
            {
                Type = StyleValues.Paragraph,
                StyleId = "Normal",
                Default = true
            };
            normalStyle.Append(new Name() { Val = "Normal" });
            normalStyle.Append(new StyleRunProperties(
                new RunFonts { Ascii = "Arial", HighAnsi = "Arial" },
                new FontSize { Val = "20" } // 10pt
            ));
            styles.Append(normalStyle);

            // Heading 1
            Style heading1 = new Style()
            {
                Type = StyleValues.Paragraph,
                StyleId = "Heading1",
                CustomStyle = false
            };
            heading1.Append(new Name() { Val = "Heading 1" });
            heading1.Append(new BasedOn() { Val = "Normal" });
            heading1.Append(new NextParagraphStyle() { Val = "Normal" });
            heading1.Append(new UIPriority() { Val = 9 });
            heading1.Append(new PrimaryStyle());
            heading1.Append(new StyleParagraphProperties(
                new KeepNext(),
                new KeepLines(),
                new SpacingBetweenLines { Before = "240", After = "60" },
                new OutlineLevel() { Val = 0 }
            ));
            heading1.Append(new StyleRunProperties(
                new Bold(),
                // new Color() { Val = "2E74B5" },
                new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" },
                new FontSize() { Val = "20" } // 10pt
            ));
            styles.Append(heading1);

            // Heading 2
            Style heading2 = new Style()
            {
                Type = StyleValues.Paragraph,
                StyleId = "Heading2",
                CustomStyle = false
            };
            heading2.Append(new Name() { Val = "Heading 2" });
            heading2.Append(new BasedOn() { Val = "Heading1" });
            heading2.Append(new NextParagraphStyle() { Val = "Normal" });
            heading2.Append(new UIPriority() { Val = 9 });
            heading2.Append(new PrimaryStyle());
            heading2.Append(new StyleParagraphProperties(
                new KeepNext(),
                new KeepLines(),
                new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" },
                new FontSize() { Val = "20" } // 10pt
            ));
            heading2.Append(new StyleRunProperties(
                new Bold(),
                // new Color() { Val = "2E74B5" },
                new RunFonts() { Ascii = "A", HighAnsi = "Calibri Light" },
                new FontSize() { Val = "26" } // 13pt
            ));
            styles.Append(heading2);

            // Heading 3
            Style heading3 = new Style()
            {
                Type = StyleValues.Paragraph,
                StyleId = "Heading3",
                CustomStyle = false
            };
            heading3.Append(new Name() { Val = "Heading 3" });
            heading3.Append(new BasedOn() { Val = "Normal" });
            heading3.Append(new NextParagraphStyle() { Val = "Heading2" });
            heading3.Append(new UIPriority() { Val = 9 });
            heading3.Append(new PrimaryStyle());
            heading3.Append(new StyleParagraphProperties(
                new KeepNext(),
                new KeepLines(),
                new SpacingBetweenLines { Before = "160", After = "20" },
                new OutlineLevel() { Val = 2 }
            ));
            heading3.Append(new StyleRunProperties(
                new Bold(),
                new Italic(),
                // new Color() { Val = "1F4D78" },
                new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" },
                new FontSize() { Val = "20" } // 10pt
            ));
            styles.Append(heading3);

            // Save styles
            stylePart.Styles = styles;
            stylePart.Styles.Save();
        }

    }
}

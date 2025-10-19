using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

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

        public void ExportHeading2Data(
            MainDocumentPart mainPart,
            YamlConfig.ReadOpenApiFile cfg)
        {
            // access
            Body? body = mainPart.Document.Body;
            if (body == null)
                return;

            // Heading
            if (cfg?.Heading2Text != null)
                body.AppendChild(CreateParagraph(
                    $"{cfg.Heading2Text}",
                    styleId: $"{_config.Heading2Style}"));

            // Intro text
            if (cfg?.Body2Text != null)
                body.AppendChild(CreateParagraph(
                    $"{cfg.Body2Text}",
                    styleId: $"{_config.BodyStyle}"));
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
                $"{opConfig?.Heading ?? _config.Heading3} {op.OperationId}",
                styleId: $"{_config.Heading3Style}"));

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

            // turn the operation's paramters into inputs
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

            // ok, suppress
            if (_config.SuppressInputs != null)
                foreach (var name in _config.SuppressInputs)
                    inputs.RemoveByName(name);

            if (_config.SuppressOutputs != null)
                foreach (var name in _config.SuppressOutputs)
                    inputs.RemoveByName(name);

            if (opConfig?.SuppressInputs != null)
                foreach (var name in opConfig.SuppressInputs)
                    inputs.RemoveByName(name);

            if (_config.SuppressOutputs != null)
                foreach (var name in _config.SuppressOutputs)
                    inputs.RemoveByName(name);

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
                tr.Append(CreateCell("Explanation ", cw[0]));
                tr.Append(CreateMergedCell($"{explanation}", true, cw[1]));
                tr.Append(CreateMergedCell("", false, cw[2]));
                tr.Append(CreateMergedCell("", false, cw[3]));
                tr.Append(CreateMergedCell("", false, cw[4]));
                table.Append(tr);
            }

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

            // 6th row: Input parameters
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

            // 7th.. row: Single input parameter
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

        static TableCell CreateCell(string text, int width, bool bold = false)
        {
            TableCell cell = new TableCell();

            TableCellProperties cellProps = new TableCellProperties(
                new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = width.ToString() },
                new TableCellBorders(
                    new TopBorder { Val = BorderValues.Single, Size = 8 },
                    new BottomBorder { Val = BorderValues.Single, Size = 8 },
                    new LeftBorder { Val = BorderValues.Single, Size = 8 },
                    new RightBorder { Val = BorderValues.Single, Size = 8 }
                ),
                new TableCellMargin(
                    new LeftMargin { Width = "80", Type = TableWidthUnitValues.Dxa },
                    new RightMargin { Width = "80", Type = TableWidthUnitValues.Dxa },
                    new TopMargin { Width = "40", Type = TableWidthUnitValues.Dxa },
                    new BottomMargin { Width = "40", Type = TableWidthUnitValues.Dxa }
                ),
                new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Top }
            );

            // Add font + size formatting
            Run run = new Run();
            RunProperties runProps = new RunProperties(
                new RunFonts { Ascii = "Arial", HighAnsi = "Arial" },
                new FontSize { Val = "16" } // 8 pt
            );

            if (bold)
                runProps.Append(new Bold());

            run.Append(runProps);
            run.Append(new Text(text ?? "") { Space = SpaceProcessingModeValues.Preserve });

            Paragraph paragraph = new Paragraph(run)
            {
                ParagraphProperties = new ParagraphProperties(new SpacingBetweenLines { After = "120" })
            };

            cell.Append(cellProps);
            cell.Append(paragraph);
            return cell;
        }

        static TableCell CreateMergedCell(string text, bool isStart, int width, bool bold = false)
        {
            TableCell cell = new TableCell();

            TableCellProperties cellProps = new TableCellProperties(
                new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = width.ToString() },
                new HorizontalMerge { Val = isStart ? MergedCellValues.Restart : MergedCellValues.Continue },
                new TableCellBorders(
                    new TopBorder { Val = BorderValues.Single, Size = 8 },
                    new BottomBorder { Val = BorderValues.Single, Size = 8 },
                    new LeftBorder { Val = BorderValues.Single, Size = 8 },
                    new RightBorder { Val = BorderValues.Single, Size = 8 }
                ),
                new TableCellMargin(
                    new LeftMargin { Width = "80", Type = TableWidthUnitValues.Dxa },
                    new RightMargin { Width = "80", Type = TableWidthUnitValues.Dxa },
                    new TopMargin { Width = "40", Type = TableWidthUnitValues.Dxa },
                    new BottomMargin { Width = "40", Type = TableWidthUnitValues.Dxa }
                ),
                new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Top }
            );

            // Add font + size formatting
            Run run = new Run();
            RunProperties runProps = new RunProperties(
                new RunFonts { Ascii = "Arial", HighAnsi = "Arial" },
                new FontSize { Val = "16" } // 8 pt
            );

            if (bold)
                runProps.Append(new Bold());

            run.Append(runProps);
            run.Append(new Text(text ?? "") { Space = SpaceProcessingModeValues.Preserve });

            Paragraph paragraph = new Paragraph(run)
            {
                ParagraphProperties = new ParagraphProperties(new SpacingBetweenLines { After = "120" })
            };

            cell.Append(cellProps);
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

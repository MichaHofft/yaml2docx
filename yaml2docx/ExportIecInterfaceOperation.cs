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
    /// This class exports little tables for the description on Interface Operations in IEC 63278-5.
    /// </summary>
    public class ExportIecInterfaceOperation
    {
#if __old

        public void Export(YamlOpenApi.OpenApiDocument doc,
            string fn)
        {
            // Create Document
            using (WordprocessingDocument wordDocument =
                WordprocessingDocument.Create(fn, WordprocessingDocumentType.Document, true))
            {
                //
                // Word init
                // see: http://www.ludovicperrichon.com/create-a-word-document-with-openxml-and-c/#table
                //

                // Add a main document part.
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();

                // Create the document structure and add some text.
                mainPart.Document = new();
                Body body = mainPart.Document.AppendChild(new Body());

                for (int ti = 0; ti < 5; ti++)
                {
                    // make a table
                    DocumentFormat.OpenXml.Wordprocessing.Table table =
                        body.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Table());

                    // new row
                    TableRow tr = table.AppendChild(new TableRow());

                    // new cell -> paragraph
                    TableCell tc = tr.AppendChild(new TableCell());
                    Paragraph para = tc.AppendChild(new Paragraph());

                    // horiz alignment
                    ParagraphProperties pp = new ParagraphProperties();
                    pp.Justification = new Justification() { Val = JustificationValues.Right };
                    para.Append(pp);

                    // vert alignment
                    var tcp = tc.AppendChild(new TableCellProperties());
                    tcp.Append(new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Bottom });

                    // .. with color
                    if (false)
                    {
                        UInt32 bgc = 0x10203040;
                        var bgs = new ColorConverter()?.ConvertToString(bgc)?.Substring(3);

                        tcp.Append(new Shading()
                        {
                            Color = "auto",
                            Fill = bgs,
                            Val = ShadingPatternValues.Clear
                        });
                    }

                    // make a run
                    var demoText = "Hallo\nWorld";
                    var run = new Run();
                    var lines = demoText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (var l in lines)
                    {
                        if (run.ChildElements.Count > 0)
                            run.AppendChild(new Break());
                        run.AppendChild(new Text(l));
                    }

                    // foreground
                    var rp = new RunProperties();
                    if (false)
                    {
                        UInt32 fgc = 0x10203040;
                        var fgs = new ColorConverter().ConvertToString(fgc)?.Substring(3);

                        rp.Append(new DocumentFormat.OpenXml.Wordprocessing.Color() { Val = fgs });
                    }

                    // further text attributes
                    if (true)
                        rp.Bold = new Bold();

                    if (true)
                        rp.Italic = new Italic();

                    if (true)
                        rp.Underline = new Underline();

                    run.RunProperties = rp;

                    // finally add to single cell
                    para.Append(run);

                    // empty rows
                    for (int i = 0; i < 3; i++)
                        body.AppendChild(new Paragraph(new Run(new Text(" "))));

                }
            }
        }
#endif 

        public void Export(YamlOpenApi.OpenApiDocument doc,
            string fn)
        {
            // Create Document
            using (WordprocessingDocument wordDoc =
                WordprocessingDocument.Create(fn, WordprocessingDocumentType.Document, true))
            {
                MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());
                Body? body = mainPart.Document.Body;
                if (body == null)
                    return;

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

                TableGrid tableGrid = new TableGrid();
                foreach (int width in cw)
                {
                    tableGrid.Append(new GridColumn() { Width = width.ToString() });
                }
                table.Append(tableGrid);

#if __old
                // Create 8 rows
                for (int r = 0; r < 8; r++)
                {
                    TableRow tr = new TableRow();

                    for (int c = 0; c < 5; c++)
                    {
                        TableCell tc;

                        // Merge columns 2–5 in rows 1 and 2
                        if ((r == 0 || r == 1) && c == 1)
                        {
                            tc = CreateMergedCell("Merged cell (row " + (r + 1) + ")", true, colWidths[c]);
                        }
                        else if ((r == 0 || r == 1) && c > 1)
                        {
                            tc = CreateMergedCell("", false, colWidths[c]);
                        }
                        else
                        {
                            tc = CreateCell($"R{r + 1}C{c + 1}", colWidths[c]);
                        }

                        tr.Append(tc);
                    }

                    table.Append(tr);
                }
#else
                
                // special col widths
                int cw04 = cw[0] + cw[1] + cw[2] + cw[3] + cw[4];
                int cw14 = cw[1] + cw[2] + cw[3] + cw[4];

                // 1st row: Header for interface operation
                {
                    TableRow tr = new TableRow();
                    tr.Append(CreateCell("Interface Operation Name ", cw[0]));
                    tr.Append(CreateMergedCell("PostSubmodelReference", true, cw[1], bold: true));
                    tr.Append(CreateMergedCell("", false, cw[2]));
                    tr.Append(CreateMergedCell("", false, cw[3]));
                    tr.Append(CreateMergedCell("", false, cw[4]));
                    table.Append(tr);
                }

                // 2nd row: Explanation
                {
                    TableRow tr = new TableRow();
                    tr.Append(CreateCell("Explanation ", cw[0]));
                    tr.Append(CreateMergedCell("Creates a Submodel Reference at the Asset Administration Shell.", true, cw[1]));
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

                // 4th row: Input data
                {
                    TableRow tr = new TableRow();
                    tr.Append(CreateMergedCell("Input Parameter", true, cw04));
                    tr.Append(CreateMergedCell("", false, cw[1]));
                    tr.Append(CreateMergedCell("", false, cw[2]));
                    tr.Append(CreateMergedCell("", false, cw[3]));
                    tr.Append(CreateMergedCell("", false, cw[4]));
                    table.Append(tr);
                }

                // 5th row: Sample input parameter
                {
                    TableRow tr = new TableRow();
                    tr.Append(CreateCell("aasId", cw[0]));
                    tr.Append(CreateCell("AssetAdministrationShell identifier.", cw[1]));
                    tr.Append(CreateCell("yes", cw[2]));
                    tr.Append(CreateCell("AssetAdministrationShellID", cw[3]));
                    tr.Append(CreateCell("1", cw[4]));
                    table.Append(tr);
                }


#endif
                body.Append(table);

                // empty rows
                for (int i = 0; i < 3; i++)
                    body.AppendChild(new Paragraph(new Run(new Text(" "))));

                // Finalize document
                mainPart.Document.Save();
            }
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
    }
}

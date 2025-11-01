using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using static Yaml2Docx.YamlConfig;

namespace Yaml2Docx
{
    public class Program
    {

        static void Main(string[] args)
        {
            // Welcome
            Console.WriteLine("Welcome to the over-engineered IEC63278-5 OpenAPI document text generator.");
            Console.WriteLine("(c) 2025 by Michael Hoffmeister, HKA");

            // Play YAML?
            //var pg = new YamlPlayground();
            //pg.Run();

            // list of already documented schema names
            var visitedSchemaExportTables = new HashSet<string>();
            var visitedSchemaExportSchema = new HashSet<string>();

            // load configuration
            var config = YamlConfig.Load(".\\configs\\yaml2docx_config.yaml");
            var wp = new ExportIecInterfaceOperation(config);

            foreach (var wfn in config.CreateWordFiles)
            {
                // Create or template
                var useTemplate = (wfn.UseTemplateFn != null);
                if (useTemplate && wfn.UseTemplateFn != null)
                {
                    Console.WriteLine($"Copying Word template file: {wfn.UseTemplateFn} to: {wfn.Fn}");
                    System.IO.File.Copy(wfn.UseTemplateFn, wfn.Fn, overwrite: true);
                }
                else
                {
                    Console.WriteLine($"Will create Word file: {wfn.Fn}");
                }

                // Create Document
                using (var wordDoc = useTemplate 
                                ? WordprocessingDocument.Open(wfn.Fn, true)
                                : WordprocessingDocument.Create(wfn.Fn, WordprocessingDocumentType.Document, true))
                {
                    MainDocumentPart? mainPart = null;
                    if (!useTemplate)
                    {
                        mainPart = wordDoc?.AddMainDocumentPart();

                        if (mainPart == null)
                        {
                            Console.WriteLine($"  ERROR: Could not create Word file body for: {wfn.Fn}");
                            continue;
                        }

                        mainPart.Document = new Document(new Body());

                        // Ensure the Styles part exists and add the default Word styles
                        if (mainPart.StyleDefinitionsPart == null)
                        {
                            var stylePart = mainPart.AddNewPart<StyleDefinitionsPart>();
                            ExportIecInterfaceOperation.GenerateDefaultStyles(stylePart);
                        }
                    } 
                    else
                    {
                        mainPart = wordDoc.MainDocumentPart;

                        if (mainPart == null)
                        {
                            Console.WriteLine($"  ERROR: Could not template Word file body for: {wfn.Fn}");
                            continue;
                        }
                    }

                    // list styles
                    if (wfn.ListStyles)
                    {
                        ExportIecInterfaceOperation.ListStyleNames(mainPart, prefix: "    ");
                    }

                    // different API files?
                    wfn.ReadOpenApiFiles = new();
                    foreach (var rof in wfn.ReadOpenApiFiles)
                    {
                        // Open
                        Console.WriteLine($"  Reading OpenAPI file: {rof.Fn}");
                        var doc = YamlOpenApi.Load(rof.Fn);
                        if (doc == null)
                        {
                            Console.WriteLine($"    ERROR: Could not read OpenAPI file: {rof.Fn}");
                            continue;
                        }

                        // List
                        if (rof.ListOperations)
                        {
                            var lst = new OpenApiLister();
                            Console.WriteLine($"    Listing pathes:");
                            lst.ListPathes(doc, prefix: "    ");
                            Console.WriteLine($"    Listing operation ids:");
                            lst.ListOperationIds(doc, prefix: "    ");
                        }

                        // over actions
                        foreach (var act in rof.Actions)
                        {
                            var actName = act.Action.Trim().ToLower();
                            if (actName == "exportpara")
                            {
                                wp.ExportParagraph(mainPart, act.ParaText, act.ParaStyle);
                            }
                            else
                            if (actName == "exporttables")
                            {
                                // Export operations
                                foreach (var opEntry in rof.UseOperations)
                                {
                                    // access for one operation, log
                                    var operationId = opEntry.Key;
                                    var opConfig = opEntry.Value;
                                    Console.WriteLine($"    Exporting operation: {operationId}");
                                    var operation = doc.FindApiOperation(operationId);
                                    if (operation == null)
                                    {
                                        Console.WriteLine($"      ERROR: Could not find operation: {operationId}");
                                        continue;
                                    }

                                    // already visited
                                    if (act.SkipIfVisited && visitedSchemaExportTables.Contains(operationId))
                                        continue;
                                    visitedSchemaExportTables.Add(operationId);

                                    // debug
                                    if (operation.OperationId == "GetAssetAdministrationShellsByQuery")
                                        ;

                                    // do
                                    wp.ExportSingleOperation(mainPart, opConfig, operation);
                                }
                            }
                            else
                            if (actName == "exportoverview")
                            {
                                // make a list of annotated operations
                                var listOfOps = new List<ExportIecInterfaceOperation.OperationTuple>();
                                int nOK = 0, nNOK = 0;
                                foreach (var opEntry in rof.UseOperations)
                                {
                                    var op = doc.FindApiOperation(opEntry.Key);
                                    if (op != null)
                                    {
                                        listOfOps.Add(new ExportIecInterfaceOperation.OperationTuple(opEntry.Value, op));
                                        nOK++;
                                    }
                                    else
                                        nNOK++;

                                }

                                // log
                                Console.WriteLine($"      Create OVERVIEW on {nOK} operations, {nNOK} not found!");

                                // do
                                wp.ExportOverviewOperation(mainPart, listOfOps);
                            }
                            else
                            if (actName == "exportyaml")
                            {
                                // Export operations
                                foreach (var opEntry in rof.UseOperations)
                                {
                                    // access for one operation, log
                                    var operationId = opEntry.Key;
                                    var opConfig = opEntry.Value;
                                    Console.WriteLine($"    Exporting operation: {operationId}");
                                    var operation = doc.FindApiOperation(operationId);
                                    if (operation == null)
                                    {
                                        Console.WriteLine($"      ERROR: Could not find operation: {operationId}");
                                        continue;
                                    }

                                    if (operation.OperationId == "GetAssetAdministrationShellsByQuery")
                                        ;

                                    // do
                                    if (act.YamlAsSource)
                                        wp.ExportSingleYamlCode(mainPart, opConfig, operation);

                                    // test do
                                    if (act.YamlAsTable)
                                        wp.ExportSingleHttpOperationDescription(doc, mainPart, opConfig, operation);
                                }
                            }
                            else
                            if (actName == "exportschema" && act.IncludeSchemas != null)
                            {
                                // try to find a schemas / data types the wished (root) schema touches
                                var schemasTouched = new HashSet<string>();
                                foreach (var sch in act.IncludeSchemas)
                                    if (sch.Trim().Length > 0)
                                        schemasTouched.Add(sch);

                                // Export the (root) schemas to find out, which other schemas are touched
                                foreach (var sch in act.IncludeSchemas)
                                    doc.RecursiveFindPropertyBundles($"#/components/schemas/{sch}", schemasTouched);

                                // the recursive search for properties might not have had a deep recursion on all
                                // schema types; therefore do it AGAIN based on the touched schemas ..
                                var toVisit = schemasTouched.ToList();
                                foreach (var sch in toVisit)
                                    doc.RecursiveFindPropertyBundles($"#/components/schemas/{sch}", schemasTouched);

                                // make a sorted list of schemas touched (no null!)
                                var schemaList = schemasTouched.Where((s) => s != null).ToList();
                                schemaList.Sort();

                                // remove if on suppressList
                                foreach (var sch in act.SuppressSchemas)
                                    if (schemaList.Contains(sch))
                                        schemaList.Remove(sch);

                                // visit them
                                foreach (var k in schemaList)
                                {
                                    // already globally visited?
                                    if (act.SkipIfVisited && visitedSchemaExportSchema.Contains(k))
                                        continue;
                                    visitedSchemaExportSchema.Add(k);

                                    // again (but not touch schemas)
                                    var pbs = doc.RecursiveFindPropertyBundles($"#/components/schemas/{k}", 
                                        schemaNotFollow: act.SchemaNotFollow);
                                    if (pbs != null)
                                    {
                                        Console.WriteLine($"    Schema to be documented: {k} .. FOUND!");
                                        wp.ExportSinglePropertyBundle(doc, mainPart, k, pbs,
                                            suppressMembers: act.SuppressMembers,
                                            schemaNotFollow: act.SchemaNotFollow);
                                    }
                                    else
                                    {
                                        Console.WriteLine($"    Schema to be documented: {k} .. missed!");
                                    }
                                }
                            }
                            else
                            if (actName == "exportpatterns")
                            {
                                // log
                                Console.WriteLine($"      Create pattern table ..");

                                // Export operations
                                wp.ExportPatternStorage(doc, mainPart);
                            }
                            else
                            {
                                // unknown action!
                                Console.WriteLine($"    ERROR: Unknown action {act.Action}!");
                                continue;
                            }
                        }
                    }

                    // grammar files
                    foreach (var rrf in wfn.ReadRailRoadFiles)
                    {
                        // Open
                        Console.WriteLine($"  Reading RailRoad file: {rrf.Fn}");
                        var rrt = new RailRoadText(rrf.Fn);
                        if (rrt == null)
                        {
                            Console.WriteLine($"    ERROR: Could not read RailRoad file: {rrf.Fn}");
                            continue;
                        }

                        // List
                        if (rrf.ListNames)
                        {
                            Console.WriteLine($"  List names of parts:");
                            foreach (var name in rrt.ListNames())
                                Console.WriteLine($"    {name}");
                        }

                        // over actions
                        foreach (var act in rrf.Actions)
                        {
                            var actName = act.Action.Trim().ToLower();
                            if (actName == "exportpara")
                            {
                                wp.ExportParagraph(mainPart, act.ParaText, act.ParaStyle);
                            }
                            else
                            if (actName == "exportrailroad")
                            {
                                // collect parts
                                var parts = new List<RailRoadText.RrPart>();
                                foreach (var pn in act.Parts)
                                {
                                    var p = rrt.FindPart(pn);
                                    if (p != null)
                                        parts.Add(p);
                                }

                                // head
                                Console.WriteLine($"  Railroad:");
                                Console.WriteLine($"");

                                // assembly
                                Console.OutputEncoding = Encoding.UTF8;
                                var assy = RailRoadText.AssembleParts(parts);
                                if (act.OutputFormat.Trim().ToLower() == "console")
                                {
                                    foreach (var ln in assy)
                                        Console.WriteLine(ln);
                                }
                            }
                        }
                    }

                    // grammar files
                    foreach (var rgf in wfn.ReadGrammarFiles)
                    {
                        // Open
                        Console.WriteLine($"  Reading grammar file: {rgf.Fn}");
                        var grammar = new GrammarText(rgf.Fn);
                        if (grammar == null)
                        {
                            Console.WriteLine($"    ERROR: Could not read grammar file: {rgf.Fn}");
                            continue;
                        }

                        // List
                        if (rgf.ListNames)
                        {
                            Console.WriteLine($"  List names of parts of grammar:");
                            foreach (var name in grammar.ListNames())
                                Console.WriteLine($"    {name}");
                        }

                        // over actions
                        foreach (var act in rgf.Actions)
                        {
                            var actName = act.Action.Trim().ToLower();
                            if (actName == "exportpara")
                            {
                                wp.ExportParagraph(mainPart, act.ParaText, act.ParaStyle);
                            }
                            else
                            if (actName == "exportgrammar")
                            {
                                // collect parts
                                var parts = new List<GrammarText.GrammarPart>();
                                foreach (var pn in act.Parts)
                                {
                                    var p = grammar.FindPart(pn);
                                    if (p != null)
                                        parts.Add(p);
                                }

                                // head
                                Console.WriteLine($"  Assembled grammar:");
                                Console.WriteLine($"");

                                // assembly
                                Console.OutputEncoding = Encoding.UTF8;
                                var assy = GrammarText.AssembleParts(parts);
                                
                                // just out?
                                if (act.OutputFormat.Trim().ToLower().Contains("console"))
                                {
                                    foreach (var ln in assy)
                                        Console.WriteLine(ln);
                                }

                                // convert via Docker to text?
                                if (act.OutputFormat.Trim().ToLower().Contains("utf8"))
                                {
                                    Console.WriteLine($"  Starting Docker {config.DockerBuildTextCmd} {config.DockerBuildTextArgs} ..");

                                    var output = new List<string>();

                                    ProcessLauncher.StartProcess(
                                        cmd: config.DockerBuildTextCmd,
                                        args: config.DockerBuildTextArgs,
                                        inputLines: assy,
                                        outputLines: output);

                                    if (output.Count < 1)
                                    {
                                        Console.WriteLine($"    ERROR: Could not generate text for input {assy.FirstOrDefault()} ..");
                                    }
                                    else
                                    {
                                        // Distance to top, Distance to bottom is automatically
                                        output.Insert(0, "");

                                        Console.WriteLine($"    Writing {output.Count()} lines to Word ..");

                                        wp.ExportMultiLineText(
                                            mainPart, act, output,
                                            fontSize: act.FontSize ?? config.GrammarCodeFontSize);
                                    }
                                }

                                // convert via Docker into an SVG?
                                if (act.OutputFormat.Trim().ToLower().Contains("svg"))
                                {
                                    Console.WriteLine($"  Starting Docker {config.DockerBuildSvgCmd} {config.DockerBuildSvgArgs} ..");

                                    ProcessLauncher.StartProcess(
                                        cmd: config.DockerBuildSvgCmd,
                                        args: config.DockerBuildSvgArgs,
                                        inputLines: assy);
                                }
                            }
                        }
                    }

                    // Finalize document
                    mainPart.Document.Save();

                }
            }            
        }
    }
}

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

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
                        ExportIecInterfaceOperation.ListStyleNames(mainPart, prefix: "  ");
                    }

                    // different API files?
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

                                    // do
                                    wp.ExportSingleYamlCode(mainPart, opConfig, operation);
                                }
                            }
                            else
                            {
                                // unknown action!
                                Console.WriteLine($"    ERROR: Unknown action {act.Action}!");
                                continue;
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

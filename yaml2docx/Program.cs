using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Yaml2Docx
{
    public class Program
    {

        static void Main(string[] args)
        {
            Console.WriteLine("Hello, World!");
            //var pg = new YamlPlayground();
            //pg.Run();

            var config = YamlConfig.Load("..\\..\\..\\..\\..\\yaml2docx_config.yaml");
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
                    if (true)
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

                        // level 2 headings
                        wp.ExportHeading2Data(mainPart, rof);

                        // Export operations
                        foreach (var opEntry in rof.ExportOperations)
                        {
                            // access, log
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

                    // Finalize document
                    mainPart.Document.Save();

                }
            }            
        }
    }
}

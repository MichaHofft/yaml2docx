
namespace Yaml2Docx
{
    public class Program
    {

        static void Main(string[] args)
        {
            Console.WriteLine("Hello, World!");
            //var pg = new YamlPlayground();
            //pg.Run();

            // var doc = YamlOpenApi.Load("..\\..\\..\\..\\..\\Plattform_i40-AssetAdministrationShellServiceSpecification-V3.1.1_SSP-001-unresolved.yaml");
            var doc = YamlOpenApi.Load("..\\..\\..\\..\\..\\Plattform_i40-AssetAdministrationShellServiceSpecification-V3.1.1_SSP-001-resolved.yaml");

            var lst = new OpenApiLister();
            
            Console.WriteLine("Listing pathes:");
            lst.ListPathes(doc);
            
            Console.WriteLine("Listing operation ids:");
            lst.ListOperationIds(doc);

            var wp = new ExportIecInterfaceOperation();
            wp.Export(doc, "..\\..\\..\\..\\..\\test2.docx");
        }
    }
}

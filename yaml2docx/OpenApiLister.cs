using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Yaml2Docx
{
    /// <summary>
    /// Contains methods to list OpenAPI elements from a YAML document.
    /// </summary>
    public class OpenApiLister
    {
        public void ListPathes(YamlOpenApi.OpenApiDocument doc, string prefix = "")
        {
            if (doc.Paths == null) return;
            foreach (var pathEntry in doc.Paths)
            {
                var pathKey = pathEntry.Key;
                var path = pathEntry.Value;
                Console.Write($"{prefix}Path: {pathKey} ");
                if (path.Get != null) Console.Write(" [GET]");
                if (path.Put != null) Console.Write(" [PUT]");
                if (path.Patch != null) Console.Write(" [PATCH]");
                if (path.Post != null) Console.Write(" [POST]");
                if (path.Delete != null) Console.Write(" [DELETE]");
                Console.WriteLine();
            }
        }

        public void ListOperationIds(YamlOpenApi.OpenApiDocument doc, string prefix = "")
        {
            if (doc.Paths == null) return;
            foreach (var pathEntry in doc.Paths)
            {
                var pathKey = pathEntry.Key;
                var path = pathEntry.Value;
                foreach (var operation in path.Operations())
                    Console.WriteLine($"{prefix}- {operation.OperationId}");
            }
        }
    }
}

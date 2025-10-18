using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;

namespace Yaml2Docx
{
    public class YamlPlayground
    {
        public class Address
        {
            public string Street { get; set; } = string.Empty;
            public string City { get; set; } = string.Empty;
            public string State { get; set; } = string.Empty;
            public string Zip { get; set; } = string.Empty;
        }

        public class Person
        {
            public string Name { get; set; } = string.Empty;
            public int Age { get; set; } = 0;
            public double HeightInInches { get; set; } = 0.0;
            public Dictionary<string, Address> Addresses { get; set; } = new();
        }

        public void Run()
        {
            var yml = @"
                name: George Washington
                age: 89
                height_in_inches: 5.75
                addresses:
                  home:
                    street: 400 Mockingbird Lane
                    city: Louaryland
                    state: Hawidaho
                    zip: 99970
            ";

            var deserializer = new DeserializerBuilder()
                .WithNamingConvention(UnderscoredNamingConvention.Instance)  // see height_in_inches in sample yml 
                .Build();

            //yml contains a string containing your YAML
            var p = deserializer.Deserialize<Person>(yml);
            var h = p.Addresses["home"];
            System.Console.WriteLine($"{p.Name} is {p.Age} years old and lives at {h.Street} in {h.City}, {h.State}.");
            // Output:
            // George Washington is 89 years old and lives at 400 Mockingbird Lane in Louaryland, Hawidaho.
        }
    }
}

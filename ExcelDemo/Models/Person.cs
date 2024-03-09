using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDemo.Models
{
    internal class Person
    {
        public int Id { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }

        public List<Person> GetSetupData()
        {
            var output = new List<Person>()
            {
                new() {Id = 1, FirstName = "Mustafa", LastName = "Fouda" },
                new() {Id = 2, FirstName = "Mohamed", LastName = "Orabi" },
            };

            return output;
        }
    }
}

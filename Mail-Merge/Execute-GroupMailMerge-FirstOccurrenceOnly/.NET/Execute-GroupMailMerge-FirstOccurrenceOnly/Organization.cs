using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Execute_GroupMailMerge_FirstOccurrenceOnly
{
    internal class Organization
    {
        public string BranchName { get; }
        public string Address { get; }
        public string Country { get; }
        public List<EmployeeDetails> Employees { get; }

        public Organization(string branchName, string address, string country, List<EmployeeDetails> employees)
        {
            BranchName = branchName;
            Address = address;
            Country = country;
            Employees = employees;
        }
    }
}

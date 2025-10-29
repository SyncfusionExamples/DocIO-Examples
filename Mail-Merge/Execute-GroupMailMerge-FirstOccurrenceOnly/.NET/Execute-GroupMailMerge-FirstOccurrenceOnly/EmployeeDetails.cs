using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Execute_GroupMailMerge_FirstOccurrenceOnly
{
    internal class EmployeeDetails
    {
        public string EmployeeName { get; }
        public string EmployeeID { get; }
        public string JoinedDate { get; }

        public EmployeeDetails(string employeeName, string employeeID, string joinedDate)
        {
            EmployeeName = employeeName;
            EmployeeID = employeeID;
            JoinedDate = joinedDate;
        }
    }
}

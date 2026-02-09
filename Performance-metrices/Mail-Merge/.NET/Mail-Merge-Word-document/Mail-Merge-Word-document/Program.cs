using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Collections;
using System.Data;
using System.Diagnostics;

class Program
{
    static void Main()
    {
        Stopwatch stopwatch = new Stopwatch();
        using (FileStream fileStream = new FileStream(Path.GetFullPath(@"../../../Data/Template.docx"), FileMode.Open, FileAccess.ReadWrite))
        {
            //Opens the template document.
            using (WordDocument document = new WordDocument(fileStream, FormatType.Docx))
            {
                //Gets the organization details as “IEnumerable” collection.
                DataSet organizationSet = BuildOrganizationsDataSet(1000, 5, 10);
                var commands = new ArrayList
                    {
                        new DictionaryEntry("Organizations", string.Empty),
                        new DictionaryEntry("Departments", "OrgId = %Organizations.OrgId%"),
                        new DictionaryEntry("Employees", "DeptId = %Departments.DeptId%")
                    };
                stopwatch.Start();
                //Performs Mail merge.
                document.MailMerge.ExecuteNestedGroup(organizationSet, commands);
                stopwatch.Stop();
                Console.WriteLine($"Time taken for mail merge in word document:" + stopwatch.Elapsed.TotalSeconds);
                //Creates file stream.
                using (FileStream outputStream = new FileStream(Path.GetFullPath(@"../../../Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    //Saves the Word document to file stream.
                    document.Save(outputStream, FormatType.Docx);
                }
            }
        }
    }
    public static DataSet BuildOrganizationsDataSet(int totalEmployees, int orgCount, int deptPerOrg)
    {
        // Get hierarchical domain data.
        var orgsInput = GetOrganizations(totalEmployees, orgCount, deptPerOrg);

        var ds = new DataSet();

        var orgs = new DataTable("Organizations");
        orgs.Columns.Add("OrgId", typeof(int));
        orgs.Columns.Add("BranchName", typeof(string));
        orgs.Columns.Add("Address", typeof(string));
        orgs.Columns.Add("City", typeof(string));
        orgs.Columns.Add("ZipCode", typeof(string));
        orgs.Columns.Add("Country", typeof(string));

        var depts = new DataTable("Departments");
        depts.Columns.Add("DeptId", typeof(int));
        depts.Columns.Add("OrgId", typeof(int));
        depts.Columns.Add("DepartmentName", typeof(string));
        depts.Columns.Add("Supervisor", typeof(string));

        var emps = new DataTable("Employees");
        emps.Columns.Add("DeptId", typeof(int));
        emps.Columns.Add("EmployeeName", typeof(string));
        emps.Columns.Add("EmployeeID", typeof(string));
        emps.Columns.Add("JoinedDate", typeof(string)); // string matches your generator

        ds.Tables.Add(orgs);
        ds.Tables.Add(depts);
        ds.Tables.Add(emps);

        // Flatten hierarchical records to rows.
        int orgId = 0, deptId = 0, empRowId = 0;
        foreach (var org in orgsInput)
        {
            orgId++;
            orgs.Rows.Add(orgId, org.BranchName, org.Address, org.City, org.ZipCode, org.Country);

            foreach (var dept in org.Departments)
            {
                deptId++;
                depts.Rows.Add(deptId, orgId, dept.DepartmentName, dept.Supervisor);

                foreach (var emp in dept.Employees)
                {
                    empRowId++;
                    emps.Rows.Add(deptId, emp.EmployeeName, emp.EmployeeID, emp.JoinedDate);
                }
            }
        }

        // Define master–detail relations that all engines can use.
        ds.Relations.Add("Org_Departments", orgs.Columns["OrgId"], depts.Columns["OrgId"]);
        ds.Relations.Add("Dept_Employees", depts.Columns["DeptId"], emps.Columns["DeptId"]);

        return ds;
    }
    public static List<Organization> GetOrganizations(int totalEmployees, int orgCount, int deptPerOrg)
    {
        var rng = new Random(42);

        int totalDepts = orgCount * deptPerOrg;
        int basePerDept = totalEmployees / totalDepts;
        int remainder = totalEmployees % totalDepts;

        string[] deptNames = {
        "Engineering","Sales","Marketing","HR","Finance",
        "Operations","Support","R&D","Logistics","Legal"
    };
        string[] firstNames = {
        "Alex","Jordan","Taylor","Sam","Casey","Jamie","Morgan","Riley","Quinn","Avery",
        "Chris","Drew","Hayden","Parker","Reese"
    };
        string[] lastNames = {
        "Smith","Johnson","Lee","Brown","Garcia","Davis","Martinez","Miller","Wilson","Clark",
        "Lopez","Young","Hall","Allen","King"
    };

        var organizations = new List<Organization>();
        int deptSequentialIndex = 0;

        for (int o = 1; o <= orgCount; o++)
        {
            var departments = new List<DepartmentDetails>();

            for (int d = 0; d < deptPerOrg; d++)
            {
                // Employee allocation for this department
                int employeesInDept = basePerDept + (remainder > 0 ? 1 : 0);
                if (remainder > 0) remainder--;

                string deptName = deptNames[d % deptNames.Length];

                // Deterministic supervisor
                string supFirst = firstNames[(o + d) % firstNames.Length];
                string supLast = lastNames[(o * 3 + d) % lastNames.Length];
                string supervisor = $"{supFirst} {supLast}";

                var employees = new List<EmployeeDetails>();

                for (int i = 0; i < employeesInDept; i++)
                {
                    string f = firstNames[rng.Next(firstNames.Length)];
                    string l = lastNames[rng.Next(lastNames.Length)];
                    string fullName = $"{f} {l}";

                    // Unique EmployeeID: EMP-<Org>-<Dept>-<Running>
                    string employeeId = $"EMP-{o:00}-{(d + 1):00}-{(deptSequentialIndex * 1000 + i + 1):0000}";

                    // Joined within last ~10 years
                    DateTime joined = DateTime.Today.AddDays(-rng.Next(60, 3650));
                    employees.Add(new EmployeeDetails(fullName, employeeId, joined.ToString("MM/dd/yyyy")));
                }

                departments.Add(new DepartmentDetails(deptName, supervisor, employees));
                deptSequentialIndex++;
            }

            // Organization info (sample address data)
            organizations.Add(new Organization(
                branchName: $"Branch {o}",
                address: $"{100 + o} Aerial Center Parkway, Suite {100 + o}",
                city: "Morrisville",
                zipcode: "27560",
                country: "USA",
                departments: departments));
        }

        return organizations;
    }
    #region Helper class
    /// <summary>
    /// Represents a class to maintain organization details.
    /// </summary>
    public class Organization
    {
        public string BranchName { get; set; }
        public string Address { get; set; }
        public string City { get; set; }
        public string ZipCode { get; set; }
        public string Country { get; set; }
        public List<DepartmentDetails> Departments { get; set; }
        public Organization(string branchName, string address, string city, string zipcode, string country, List<DepartmentDetails> departments)
        {
            BranchName = branchName;
            Address = address;
            City = city;
            ZipCode = zipcode;
            Country = country;
            Departments = departments;
        }
    }
    /// <summary>
    /// Represents a class to maintain department details.
    /// </summary>
    public class DepartmentDetails
    {
        public string DepartmentName { get; set; }
        public string Supervisor { get; set; }
        public List<EmployeeDetails> Employees { get; set; }
        public DepartmentDetails(string departmentName, string supervisor, List<EmployeeDetails> employees)
        {
            DepartmentName = departmentName;
            Supervisor = supervisor;
            Employees = employees;
        }
    }
    /// <summary>
    /// Represents a class to maintain employee details.
    /// </summary>
    public class EmployeeDetails
    {
        public string EmployeeName { get; set; }
        public string EmployeeID { get; set; }
        public string JoinedDate { get; set; }
        public EmployeeDetails(string employeeName, string employeeID, string joinedDate)
        {
            EmployeeName = employeeName;
            EmployeeID = employeeID;
            JoinedDate = joinedDate;
        }
    }
    #endregion
}

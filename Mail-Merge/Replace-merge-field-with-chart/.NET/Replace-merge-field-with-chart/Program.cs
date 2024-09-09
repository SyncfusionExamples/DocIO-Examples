using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.OfficeChart;
using System.Collections.Generic;
using System.IO;

namespace Replace_merge_field_with_chart
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.ReadWrite))
            {
                //Loads an existing Word document into DocIO instance.
                using (WordDocument document = new WordDocument(fileStream, FormatType.Automatic))
                {
                    //Gets the employee details as “IEnumerable” collection
                    List<Employees> employeeList = GetEmployeeData();
                    //Creates an instance of MailMergeDataTable by specifying MailMerge group name and IEnumerable collection.
                    MailMergeDataTable dataTable = new MailMergeDataTable("Employees", employeeList);
                    //Uses the mail merge event handler to insert chart during mail merge.
                    document.MailMerge.MergeField += new MergeFieldEventHandler(MergeField_EmployeeGraph);
                    //Performs Mail merge.
                    document.MailMerge.ExecuteGroup(dataTable);
                    //Unhooks the event after mail merge execution.
                    document.MailMerge.MergeField -= new MergeFieldEventHandler(MergeField_EmployeeGraph);
                    //Creates file stream.
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            }          
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(Path.GetFullPath(@"Output/Result.docx")) { UseShellExecute = true });
        }
        #region Helper methods
        /// <summary>
        /// Gets the employee data to perform mail merge. 
        /// </summary>
        /// <returns></returns>
        public static List<Employees> GetEmployeeData()
        {           
            //Creates graph data for first employee.
            List<object[]> graphDetailsForEmployee1 = new List<object[]>();
            graphDetailsForEmployee1.Add(new object[] { "Month", "Highest Sale", "Average Sale", "Lowest Sale" });
            graphDetailsForEmployee1.Add(new object[] { "September", 67, 55, 12 });
            graphDetailsForEmployee1.Add(new object[] { "October", 74, 71, 70 });
            graphDetailsForEmployee1.Add(new object[] { "November", 81, 74, 60 });
            graphDetailsForEmployee1.Add(new object[] { "December", 96, 71, 20});

            //Creates graph data for the second employee.
            List<object[]> graphDetailsForEmployee2 = new List<object[]>();
            graphDetailsForEmployee2.Add(new object[] { "Month", "Highest Sale", "Average Sale", "Lowest Sale" });
            graphDetailsForEmployee2.Add(new object[] { "September", 100, 65, 50 });
            graphDetailsForEmployee2.Add(new object[] { "October", 72, 34, 15 });
            graphDetailsForEmployee2.Add(new object[] { "November", 150, 81, 63 });
            graphDetailsForEmployee2.Add(new object[] { "December", 91, 75, 50 });

            //Creates graph data for the third employee.
            List<object[]> graphDetailsForEmployee3 = new List<object[]>();
            graphDetailsForEmployee3.Add(new object[] { "Month", "Highest Sale", "Average Sale", "Lowest Sale" });
            graphDetailsForEmployee3.Add(new object[] { "September", 58, 26, 14 });
            graphDetailsForEmployee3.Add(new object[] { "October", 55, 45, 30 });
            graphDetailsForEmployee3.Add(new object[] { "November", 62, 51, 23 });
            graphDetailsForEmployee3.Add(new object[] { "December", 72, 45, 11 });

            //Adds all details in employee data collection for all employees.
            List<Employees> employeeData = new List<Employees>();
            employeeData.Add(new Employees("Nancy", "Davolio", "1", "505 - 20th Ave. E. Apt. 2A,", "Seattle", "USA", graphDetailsForEmployee1));
            employeeData.Add(new Employees("Andrew", "Fuller", "2", "908 W. Capital Way", "Tacoma", "USA", graphDetailsForEmployee2));
            employeeData.Add(new Employees("Margaret", "Peacock", "3", "4110 Old Redmond Rd.", "Redmond", "USA", graphDetailsForEmployee3));
            return employeeData;
        }
        /// <summary>
        /// Represents the method that handles MergeField event.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        private static void MergeField_EmployeeGraph(object sender, MergeFieldEventArgs args)
        {
            //Creates chart based on the field value and insert into the Word document.
            if (args.FieldName == "GraphDetails")
            {
                //Gets their owner row.
                WParagraph paragraph = args.CurrentMergeField.OwnerParagraph;
                //Gets the field value.
                List<object[]> graphDetails = args.FieldValue as List<object[]>;
                //Creates the chart.
                WChart chart = CreateChart(paragraph.Document, graphDetails);
                int indexOfField = paragraph.ChildEntities.IndexOf(args.CurrentMergeField);
                //Clears the field;
                args.Text = string.Empty;
                //Inserts the chart at corresponding field location.
                paragraph.ChildEntities.Insert(indexOfField, chart);
            }
        }
        /// <summary>
        /// Creates the chart based on the graph data.
        /// </summary>
        /// <param name="document"></param>
        private static WChart CreateChart(WordDocument document, List<object[]> graphDetails)
        {
            //Create the new chart.
            WChart chart = new WChart(document);
            chart.Width = 410;
            chart.Height = 250;
            chart.ChartType = OfficeChartType.Column_Clustered;
            //Assign data.
            AddChartData(chart, graphDetails);
            //Set a chart title.
            chart.ChartTitle = "Sales Report";
            //Set Datalabels.
            IOfficeChartSerie serie1 = chart.Series.Add("Highest Sale");
            //Set the data range of chart series – start row, start column, end row and end column.
            serie1.Values = chart.ChartData[2, 2, 5, 2];
            IOfficeChartSerie serie2 = chart.Series.Add("Average Sale");
            //Set the data range of chart series start row, start column, end row and end column.
            serie2.Values = chart.ChartData[2, 3, 5, 3];
            IOfficeChartSerie serie3 = chart.Series.Add("Lowest Sale");
            //Set the data range of chart series start row, start column, end row and end column.
            serie3.Values = chart.ChartData[2, 4, 5, 4];
            //Set the data range of the category axis.
            chart.PrimaryCategoryAxis.CategoryLabels = chart.ChartData[2, 1, 6, 1];
            //Set legend.
            chart.HasLegend = true;
            chart.Legend.Position = OfficeLegendPosition.Bottom;
            //Hiding major gridlines
            chart.PrimaryValueAxis.HasMajorGridLines = false;
            return chart;
        }
        /// <summary>
        /// Set the values for the chart.
        /// </summary>
        private static void AddChartData(WChart chart, List<object[]> graphDetails)
        {
            //Set the value for chart data.
            int rowIndex = 1;
            int colIndex = 1;
            //Get the value from the DataTable and set the value for chart data
            foreach (object[] row in graphDetails)
            {
                foreach (object value in row)
                {
                    chart.ChartData.SetValue(rowIndex, colIndex, value);
                    colIndex++;
                    if (colIndex == 5)
                        break;
                }
                colIndex = 1;
                rowIndex++;
            }
        }
        #endregion
    }
    #region Helper Class
    /// <summary>
    /// Represents a class to maintain employee details.
    /// </summary>
    public class Employees
    {
            public string FirstName { get; set; }
            public string LastName { get; set; }
            public string EmployeeID { get; set; }
            public string Address { get; set; }
            public string City { get; set; }    
            public string Country { get; set; }
            public List<object[]> GraphDetails { get; set; }
            public Employees(string firstName, string lastName, string employeeID, string address, string city, string country, List<object[]> graphDetails)
            {
                FirstName = firstName;
                LastName = lastName;
                EmployeeID = employeeID;
                Address = address;
                City = city;              
                Country = country;
                GraphDetails = graphDetails;               
            }
    }
    #endregion
}

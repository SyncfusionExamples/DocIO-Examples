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
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"../../../Template.docx"), FileMode.Open, FileAccess.ReadWrite))
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
                    //Creates file stream.
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            }
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(Path.GetFullPath(@"../../../Result.docx")) { UseShellExecute = true });
        }

        #region Helper methods
        /// <summary>
        /// Gets the employee data to perform mail merge. 
        /// </summary>
        /// <returns></returns>
        public static List<Employees> GetEmployeeData()
        {
            
            //Creates graph data for first employee.
            List<string[]> graphDetails = new List<string[]>();
            graphDetails.Add(new string[] { "Month", "Highest Sale", "Average Sale", "Lowest Sale" });
            graphDetails.Add(new string[] { "September", "67", "55", "12" });
            graphDetails.Add(new string[] { "October", "74", "71", "70" });
            graphDetails.Add(new string[] { "November", "81", "74", "60" });
            graphDetails.Add(new string[] { "December", "96", "71", "20"});

            //Creates graph data for first employee.
            List<string[]> graphDetails2 = new List<string[]>();
            graphDetails2.Add(new string[] { "Month", "Highest Sale", "Average Sale", "Lowest Sale" });
            graphDetails2.Add(new string[] { "September", "100", "65", "50" });
            graphDetails2.Add(new string[] { "October", "72", "34", "15" });
            graphDetails2.Add(new string[] { "November", "150", "81", "63" });
            graphDetails2.Add(new string[] { "December", "91", "75", "50" });

            //Creates graph data for first employee.
            List<string[]> graphDetails3 = new List<string[]>();
            graphDetails3.Add(new string[] { "Month", "Highest Sale", "Average Sale", "Lowest Sale" });
            graphDetails3.Add(new string[] { "September", "58", "26", "14" });
            graphDetails3.Add(new string[] { "October", "55", "45", "30" });
            graphDetails3.Add(new string[] { "November", "62", "51", "23" });
            graphDetails3.Add(new string[] { "December", "72", "45", "11" });

            //Adds all details in employee data collection for all employees.
            List<Employees> employeeData = new List<Employees>();
            employeeData.Add(new Employees("Nancy", "Davolio", "1", "505 - 20th Ave. E. Apt. 2A,", "Seattle", "USA", graphDetails));
            employeeData.Add(new Employees("Andrew", "Fuller", "2", "908 W. Capital Way", "Tacoma", "USA", graphDetails2));
            employeeData.Add(new Employees("Margaret", "Peacock", "3", "4110 Old Redmond Rd.", "Redmond", "USA", graphDetails3));

            

            return employeeData;
        }
        /// <summary>
        /// 
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
                List<string[]> graphDetails = args.FieldValue as List<string[]>;
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
        private static WChart CreateChart(WordDocument document, List<string[]> graphDetails)
        {

            //Create the new chart.
            WChart chart = new WChart(document);
            chart.Width = 446;
            chart.Height = 150;
            chart.ChartType = OfficeChartType.Column_Clustered;
            //Assign data.
            AddChartData(chart, graphDetails);
            //Set a chart title.
            chart.ChartTitle = "Sales Report";
            //Set Datalabels.
            IOfficeChartSerie serie1 = chart.Series.Add("Highest Mark");
            //Set the data range of chart series – start row, start column, end row and end column.
            serie1.Values = chart.ChartData[2, 2, 5, 2];
            IOfficeChartSerie serie2 = chart.Series.Add("Average Mark");
            //Set the data range of chart series start row, start column, end row and end column.
            serie2.Values = chart.ChartData[2, 3, 5, 3];
            IOfficeChartSerie serie3 = chart.Series.Add("Mark");
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
        private static void AddChartData(WChart chart, List<string[]> graphDetails)
        {
            //Set the value for chart data.
            int rowIndex = 1;
            int colIndex = 1;
            //Get the value from the DataTable and set the value for chart data
            foreach (string[] row in graphDetails)
            {
                foreach (string val in row)
                {
                    string value = val.ToString();
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

    public class Employees
        {
            public string FirstName { get; set; }
            public string LastName { get; set; }
            public string EmployeeID { get; set; }
            public string Address { get; set; }
            public string City { get; set; }    
            public string Country { get; set; }
            public List<string[]> GraphDetails { get; set; }
            public Employees(string firstName, string lastName, string employeeID, string address, string city, string country, List<string[]> graphDetails)
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
}

using Syncfusion.DocIO.DLS;
using Syncfusion.OfficeChart;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;

namespace Create_Pie_chart_from_database
{
    class Program
    {
        static void Main(string[] args)
        {
            //Create a new instance of WordDocument.
            using (WordDocument document = new WordDocument())
            {
                document.EnsureMinimal();
                //Get the data table
                DataTable dataTable = GetDataTable();
                //Create and append the chart to the paragraph.
                WChart chart = document.LastParagraph.AppendChart(446, 270);
                chart.ChartType = OfficeChartType.Pie;
                //Assign the data.
                AddChartData(chart, dataTable);
                //Set a chart title.
                chart.ChartTitle = "Best Selling Products";
                IOfficeChartSerie pieSeries = chart.Series.Add("Sales");
                pieSeries.Values = chart.ChartData[2, 2, 11, 2];
                //Set the data label.
                pieSeries.DataPoints.DefaultDataPoint.DataLabels.IsValue = true;
                pieSeries.DataPoints.DefaultDataPoint.DataLabels.Position = OfficeDataLabelPosition.Outside;
                //Set the category labels.
                chart.PrimaryCategoryAxis.CategoryLabels = chart.ChartData[2, 1, 11, 1];
                //Set the legend.
                chart.HasLegend = true;
                //Save the Word document.
                document.Save(Path.GetFullPath(@"../../Result.docx"));
            }        
        }

        #region Helper methods
        /// <summary>
        /// Get the data to create  pie chart.
        /// </summary>
        private static DataTable GetDataTable()
        {
            string path = Path.GetFullPath(@"../../Data/DataBase.mdb");
            //Create a new instance of OleDbConnection
            OleDbConnection connection = new OleDbConnection();
            //Set the string to open a Database
            connection.ConnectionString = "Provider=Microsoft.JET.OLEDB.4.0;Password=\"\";User ID=Admin;Data Source=" + path;
            //Open the Database connection
            connection.Open();
            //Get all the data from the Database
            OleDbCommand query = new OleDbCommand("select * from Products", connection);
            //Create a new instance of OleDbDataAdapter
            OleDbDataAdapter adapter = new OleDbDataAdapter(query);
            //Create a new instance of DataSet
            DataSet dataSet = new DataSet();
            //Add rows in the Dataset
            adapter.Fill(dataSet);
            //Create a DataTable from the Dataset
            DataTable table = dataSet.Tables[0];
            table.TableName = "Products";
            return table;
        }
        /// <summary>
        /// Set the value for the chart.
        /// </summary>
        private static void AddChartData(WChart chart, DataTable dataTable)
        {
            //Set the value for chart data.
            chart.ChartData.SetValue(1, 1, "Names");
            chart.ChartData.SetValue(1, 2, "Product");

            int rowIndex = 2;
            int colIndex = 1;
            //Get the value from the DataTable and set the value for chart data
            foreach (DataRow row in dataTable.Rows)
            {
                foreach (object val in row.ItemArray)
                {
                    string value = val.ToString();
                    chart.ChartData.SetValue(rowIndex, colIndex, value);
                    colIndex++;
                    if (colIndex == 3)
                        break;
                }
                colIndex = 1;
                rowIndex++;
            }
        }
        #endregion
    }
}
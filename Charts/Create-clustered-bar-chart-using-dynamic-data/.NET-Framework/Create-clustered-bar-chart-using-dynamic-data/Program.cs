using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using Syncfusion.OfficeChart;
using System.Data;
using System.IO;
using System.Data.OleDb;

namespace Create_clustered_bar_chart_using_dynamic_data
{
    internal class Program
    {
        static void Main(string[] args)
        {
            //Creates a new WordDocument.
            using (WordDocument document = new WordDocument())
            {
                //Adds section to the document.
                IWSection sec = document.AddSection();
                //Adds paragraph to the section.
                IWParagraph paragraph = sec.AddParagraph();
                //Creates and appends chart to the paragraph.
                WChart chart = paragraph.AppendChart(470, 300);

                //Sets chart type.
                chart.ChartType = OfficeChartType.Bar_Clustered;

                //Assign data range.
                chart.DataRange = chart.ChartData[1, 1, 6, 4];
                chart.IsSeriesInRows = false;

                //Gets the data table from the database.
                DataTable dataTable = GetDataTable();
                //Sets data to the chart - RowIndex, columnIndex and data.
                SetChartData(chart, dataTable);

                //Apply chart elements.
                //Set chart title.
                chart.ChartTitle = "Clustered Bar Chart";

                //Sets Datalabels.
                IOfficeChartSerie serie1 = chart.Series[0];
                IOfficeChartSerie serie2 = chart.Series[1];
                IOfficeChartSerie serie3 = chart.Series[2];

                serie1.DataPoints.DefaultDataPoint.DataLabels.IsValue = true;
                serie2.DataPoints.DefaultDataPoint.DataLabels.IsValue = true;
                serie3.DataPoints.DefaultDataPoint.DataLabels.IsValue = true;
                serie1.DataPoints.DefaultDataPoint.DataLabels.Position = OfficeDataLabelPosition.Center;
                serie2.DataPoints.DefaultDataPoint.DataLabels.Position = OfficeDataLabelPosition.Center;
                serie3.DataPoints.DefaultDataPoint.DataLabels.Position = OfficeDataLabelPosition.Center;

                //Sets legend.
                chart.HasLegend = true;
                chart.Legend.Position = OfficeLegendPosition.Bottom;
                //Creates file stream.
                using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    //Saves the Word document to file stream.
                    document.Save(outputFileStream, FormatType.Docx);
                }
            }
            System.Diagnostics.Process.Start(Path.GetFullPath(@"../../Result.docx"));
        }
        /// <summary>
        /// Gets the data for chart.
        /// </summary>
        private static DataTable GetDataTable()
        {
            string path = Path.GetFullPath(@"../../Data/DataBase.mdb");
            //Create a new instance of OleDbConnection
            OleDbConnection connection = new OleDbConnection();
            //Sets the string to open a Database
            connection.ConnectionString = "Provider=Microsoft.JET.OLEDB.4.0;Password=\"\";User ID=Admin;Data Source=" + path;
            //Opens the Database connection
            connection.Open();
            //Get all the data from the Database
            OleDbCommand query = new OleDbCommand("select * from Fruits", connection);
            //Create a new instance of OleDbDataAdapter
            OleDbDataAdapter adapter = new OleDbDataAdapter(query);
            //Create a new instance of DataSet
            DataSet dataSet = new DataSet();
            //Adds rows in the Dataset
            adapter.Fill(dataSet);
            //Create a DataTable from the Dataset
            DataTable table = dataSet.Tables[0];
            return table;
        }
        /// <summary>
        /// Set the values for the chart.
        /// </summary>
        private static void SetChartData(WChart chart, DataTable dataTable)
        {
            //Sets the heading for chart data.
            chart.ChartData.SetValue(1, 1, "Fruits");
            chart.ChartData.SetValue(1, 2, "Joey");
            chart.ChartData.SetValue(1, 3, "Mathew");
            chart.ChartData.SetValue(1, 4, "Peter");

            int rowIndex = 2;
            int colIndex = 1;
            //Get the values from the DataTable and set the value for chart data.
            foreach (DataRow row in dataTable.Rows)
            {
                foreach (object val in row.ItemArray)
                {
                    string value = val.ToString();
                    //Sets data to the chart - RowIndex, columnIndex and data.
                    chart.ChartData.SetValue(rowIndex, colIndex, value);
                    colIndex++;
                    if (colIndex == (row.ItemArray.Length + 1))
                        break;
                }
                colIndex = 1;
                rowIndex++;
            }
        }
    }
}

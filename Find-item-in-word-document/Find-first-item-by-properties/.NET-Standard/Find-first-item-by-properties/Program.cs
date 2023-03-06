using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.OfficeChart;
using System.IO;

namespace Find_first_item_by_properties
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"../../../Input.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Opens an existing Word document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic))
                {
                    string[] propertyNames = { "ChartType", "ChartTitle" };
                    string[] propertyValues = { OfficeChartType.Pie.ToString(), "Sales" };
                    //Find the chart by ChartType and ChartTitle.
                    WChart chart = document.FindItemByProperties(EntityType.Chart, propertyNames, propertyValues) as WChart;
                    //Rename the ChartTitle.
                    if (chart != null)
                        chart.ChartTitle = "Sales Analysis";

                    propertyNames = new string[] { "Title", "Rows.Count" };
                    propertyValues = new string[] { "SupplierDetails", "6" };
                    //Find the table by Title and Rows Count
                    WTable table = document.FindItemByProperties(EntityType.Table, propertyNames, propertyValues) as WTable;
                    //Remove the table in document.
                    if (table != null)
                        table.OwnerTextBody.ChildEntities.Remove(table);
                    //Creates file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Sample.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}

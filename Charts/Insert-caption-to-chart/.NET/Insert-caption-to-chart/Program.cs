using System.IO;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.OfficeChart;

namespace Insert_caption_to_chart
{
    class Program
    {
        static void Main(string[] args)
        {
            Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("NxYtFisQPR08Cit/VkR+XU9FfV5AQmBIYVp/TGpJfl96cVxMZVVBJAtUQF1hTH9SdENiWHtZc3ZVRWFeWkd1");
            // Load existing Word document.
            using (FileStream inputFileStream = new FileStream(Path.GetFullPath(@"../../../Data/Template.docx"), FileMode.Open, FileAccess.ReadWrite))
            {
                // Initialize the Word document with the input file stream.
                using (WordDocument document = new WordDocument(inputFileStream, FormatType.Automatic))
                {
                    Entity entity = document.FindItemByProperty(EntityType.Chart, null, null);
                    WChart chart = entity as WChart;
                   if (chart != null)
                    {
                        //Mention caption text here.
                        string captionName = "Chart";
                        //Add caption to the chart.
                        AddCaptionToChart(chart, captionName, CaptionNumberingFormat.Number, CaptionPosition.AfterImage);
                        //Update fields in the Word document.
                        document.UpdateDocumentFields();
                    }
                    //Create a file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Save the Word document to the file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
        /// <summary>
        ///  Add caption to the chart.
        /// </summary>
        public static void AddCaptionToChart(WChart chart, string captionName, CaptionNumberingFormat format, CaptionPosition captionPosition)
        {
            IWParagraph ownerParagraph = chart.OwnerParagraph;
            WTextBody body = ownerParagraph.Owner as WTextBody;
            WParagraph paragraph = null;
            if (body != null)
            {
                //Get the index of the owner paragraph.
                int index = GetIndexInOwnerCollection(ownerParagraph); 
                paragraph = new WParagraph(chart.Document);
                paragraph.AppendText(captionName + " ");
                captionName = captionName.Replace(" ", "_");
                paragraph.ApplyStyle(BuiltinStyle.Caption);
                WSeqField field = (WSeqField)paragraph.AppendField(captionName, FieldType.FieldSequence);
                field.NumberFormat = format;
                int chartIndex = ownerParagraph.Items.IndexOf(chart);

                // Set needed formatting and paragraph location dependently on captionPosition value
                if (captionPosition == CaptionPosition.AfterImage)
                {
                    ownerParagraph.ParagraphFormat.KeepFollow = true;
                    body.ChildEntities.Insert(index + 1, paragraph);
                }
                else
                {
                    paragraph.ParagraphFormat.KeepFollow = true;
                    int captionIndex = (chartIndex == 0) ? index : index + 1;

                    body.ChildEntities.Insert(captionIndex, paragraph);

                    if (chartIndex > 0)
                    {
                        ownerParagraph.Items.RemoveAt(chartIndex);
                        WParagraph newParagraph = new WParagraph(chart.Document);
                        newParagraph.Items.Insert(0, chart);
                        body.ChildEntities.Insert(captionIndex + 1, newParagraph);
                    }
                }
                ApplyFormattingForCaption(paragraph);
            }
        }
        /// <summary>
        /// This methode is used to get the index of the paragraph.
        /// </summary>
        public static int GetIndexInOwnerCollection(IWParagraph ownerParagraph)
        {
         
            ICompositeEntity composite = ownerParagraph.Owner as ICompositeEntity;

            if (composite != null)
            {
                return composite.ChildEntities.IndexOf(ownerParagraph);
            }
            //If item is inside inline content control.
            else if (ownerParagraph is InlineContentControl)
                return (ownerParagraph as InlineContentControl).ParagraphItems.IndexOf(ownerParagraph);

            return -1;
        }
        /// <summary>
        /// Apply formattings for image caption paragraph
        /// </summary>
        public static void ApplyFormattingForCaption(WParagraph paragraph)
        {
            //Align the caption
            paragraph.ParagraphFormat.HorizontalAlignment = Syncfusion.DocIO.DLS.HorizontalAlignment.Center;
            //Sets after spacing
            paragraph.ParagraphFormat.AfterSpacing = 1.5f;
            //Sets before spacing
            paragraph.ParagraphFormat.BeforeSpacing = 1.5f;
        }
    }
}

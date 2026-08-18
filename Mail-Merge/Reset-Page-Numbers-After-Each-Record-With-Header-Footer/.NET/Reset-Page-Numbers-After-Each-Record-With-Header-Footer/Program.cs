using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Dynamic;
using System.IO;
using System.Linq;

namespace Reset_Page_Numbers_After_Each_Record_With_Header_Footer
{
    class Program
    {
        static void Main(string[] args)
        {
            // Open document with using file stream.
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"../../../Data/Template.docx"), FileMode.Open, FileAccess.ReadWrite))
            {
                //Opens the template document.
                using (WordDocument document = new WordDocument(fileStream, FormatType.Docx))
                {
                    #region Execute mail merge
                    //Creates an instance of the MailMergeDataSet.
                    MailMergeDataSet dataSet = new MailMergeDataSet();
                    //Creates the mail merge data table in order to perform mail merge.
                    MailMergeDataTable dataTable = new MailMergeDataTable("Customers", GetCustomers());
                    dataSet.Add(dataTable);
                    dataTable = new MailMergeDataTable("Orders", GetOrders());
                    dataSet.Add(dataTable);
                    List<DictionaryEntry> commands = new List<DictionaryEntry>();
                    //DictionaryEntry contain "Source table" (key) and "Command" (value).
                    DictionaryEntry entry = new DictionaryEntry("Customers", string.Empty);
                    commands.Add(entry);
                    //Retrieves the customer details.
                    entry = new DictionaryEntry("Orders", "CustomerID = %Customers.CustomerID%");
                    commands.Add(entry);
                    //Performs the mail merge operation with the dynamic collection.
                    document.MailMerge.ExecuteNestedGroup(dataSet, commands);
                    #endregion

                    #region Split Word document by sections
                    //Find all the occurance of place holder - #SectionBreak#
                    Syncfusion.DocIO.DLS.TextSelection[] selections = document.FindAll("#SectionBreak#", false, false);
                    //Loop through all the text selection
                    for (int i = selections.Length - 1; i >= 0; i--)
                    {
                        Syncfusion.DocIO.DLS.TextSelection selection = selections[i];
                        //Get the owner paragraph of the selected text
                        WParagraph para = selection.GetAsOneRange().OwnerParagraph;
                        WSection srcSection = document.LastSection;
                        //Insert section break
                        InsertSectionBreak(para, srcSection);
                        WSection curSection = GetSection(para);
                        //Removes the place holder.
                        curSection.Body.ChildEntities.Remove(para);
                    }
                    #endregion

                    #region Resets the page number
                    //Iterates each section from Word document.
                    foreach (WSection section in document.Sections)
                    {
                        //Resets the page number.
                        section.PageSetup.RestartPageNumbering = true;
                        section.PageSetup.PageStartingNumber = 1;
                    }
                    //Updates fields in Word document.
                    document.UpdateDocumentFields(true);
                    #endregion

                    //Creates file stream.
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            }
        }

        #region Helper Methods
        /// <summary>
        /// Insert Section break code
        /// </summary>
        /// <param name="bodyItem"></param>
        /// <param name="breakCode"></param>
        /// <returns></returns>
        private static WSection InsertSectionBreak(TextBodyItem bodyItem, WSection srcSection)
        {
            //Get the current section of the body item
            var currentSection = GetSection(bodyItem);

            // Identify all the items in the same Section that are positioned after the bodyItem. These body items need to be cut and pasted to the new section.
            int numBodyItemsToStay = GetIndex(bodyItem) + 1;
            var entityCollection = currentSection.Body.ChildEntities;
            var bodyItemsToMove = entityCollection.Cast<TextBodyItem>()
                                              .Skip(numBodyItemsToStay)
                                              .ToList();

            //Create a new section that is positioned after the current section.
            var newSection = new WSection(bodyItem.Document);
            //Updates the section properties of the template document.
            CopySectionProperties(newSection, srcSection);

            //Copy headers and footers from current section into the new section
            CopyHeaderAndFooter(newSection, currentSection, false);
            //Add new section as a sibling of current section
            AddSiblings(currentSection, new[] { newSection });

            // Cut and paste each marked body item from the current section to the new section.
            foreach (var bodyItemToMove in bodyItemsToMove)
            {
                newSection.Body.ChildEntities.Add(bodyItemToMove);
            }
            return newSection;
        }
        /// <summary>
        /// Get the index of the particular entity
        /// </summary>
        /// <param name="entity"></param>
        /// <returns></returns>
        private static int GetIndex(IEntity entity)
        {
            ICompositeEntity container = entity.Owner as ICompositeEntity;
            if (container == null)
            {
                throw new ApplicationException("Entity is not index-able as it does not have a valid container.");
            }

            return container.ChildEntities.IndexOf(entity);
        }
        /// <summary>
        /// Geth the section of the specified entity
        /// </summary>
        /// <param name="entity"></param>
        /// <returns></returns>
        private static WSection GetSection(IEntity entity)
        {
            if (entity is WSection)
            {
                return (WSection)entity;
            }

            if (entity is WordDocument)
            {
                throw new ApplicationException("WordDocument does not belong to any sections.");
            }

            // Traverse the tree bottom-up until the Section is found.
            IEntity parentEntity = entity.Owner;
            while (parentEntity != null)
            {
                if (parentEntity is WSection)
                {
                    return (WSection)parentEntity;
                }

                parentEntity = parentEntity.Owner;
            }

            // Unable to find the Section this entity belongs to. This entity is most likely not attached to any containers yet.
            return null;
        }
        /// <summary>
        /// Copy headers and footers
        /// </summary>
        /// <param name="section"></param>
        /// <param name="sectionToCopyFrom"></param>
        /// <param name="copyLinkToPrevious"></param>
        /// <param name="copyFirstPageHeaderFooters"></param>
        private static void CopyHeaderAndFooter(WSection destSection,
                                             WSection srcSection,
                                             bool copyFirstPageHeaderFooters)
        {
            //Copy child entity of Headers and footers
            CopyFrom(destSection.HeadersFooters.EvenFooter, srcSection.HeadersFooters.EvenFooter);
            CopyFrom(destSection.HeadersFooters.EvenHeader, srcSection.HeadersFooters.EvenHeader);
            CopyFrom(destSection.HeadersFooters.Header, srcSection.HeadersFooters.Header);
            CopyFrom(destSection.HeadersFooters.Footer, srcSection.HeadersFooters.Footer);
            CopyFrom(destSection.HeadersFooters.OddFooter, srcSection.HeadersFooters.OddFooter);
            CopyFrom(destSection.HeadersFooters.OddHeader, srcSection.HeadersFooters.OddHeader);

            //Copy first page header and footer
            if (copyFirstPageHeaderFooters)
            {
                CopyFrom(destSection.HeadersFooters.FirstPageFooter, srcSection.HeadersFooters.FirstPageFooter);
                CopyFrom(destSection.HeadersFooters.FirstPageHeader, srcSection.HeadersFooters.FirstPageHeader);
            }
        }
        /// <summary>
        /// Copy child entity
        /// </summary>
        /// <param name="textBody"></param>
        /// <param name="otherTextBody"></param>
		private static void CopyFrom(WTextBody destTextBody, WTextBody srcTextBody)
        {
            destTextBody.ChildEntities.Clear();
            //Loop through all the child entites of textbody
            foreach (Entity childEntity in srcTextBody.ChildEntities)
            {
                //Clone and add the child entity to the new text body
                destTextBody.ChildEntities.Add(childEntity.Clone());
            }
        }
        /// <summary>
        /// Copy page setup
        /// </summary>
        /// <param name="section"></param>
        /// <param name="sectionToCopyFrom"></param>
        private static void CopySectionProperties(IWSection newSection, IWSection srcSection)
        {
            //Updates section break code.
            newSection.BreakCode = srcSection.BreakCode;
            //Updates column size.
            foreach (Column column in srcSection.Columns)
            {
                newSection.AddColumn(column.Width, column.Space);
            }
            //Updates section page set up.
            newSection.PageSetup.Bidi = srcSection.PageSetup.Bidi;

            newSection.PageSetup.DifferentFirstPage = srcSection.PageSetup.DifferentFirstPage;
            newSection.PageSetup.DifferentOddAndEvenPages = srcSection.PageSetup.DifferentOddAndEvenPages;
            newSection.PageSetup.FooterDistance = srcSection.PageSetup.FooterDistance;
            newSection.PageSetup.HeaderDistance = srcSection.PageSetup.HeaderDistance;
            newSection.PageSetup.IsFrontPageBorder = srcSection.PageSetup.IsFrontPageBorder;
            newSection.PageSetup.Margins = srcSection.PageSetup.Margins;
            newSection.PageSetup.Orientation = srcSection.PageSetup.Orientation;
            newSection.PageSetup.PageBorderOffsetFrom = srcSection.PageSetup.PageBorderOffsetFrom;
            newSection.PageSetup.PageBordersApplyType = srcSection.PageSetup.PageBordersApplyType;
            newSection.PageSetup.PageNumberStyle = srcSection.PageSetup.PageNumberStyle;
            newSection.PageSetup.PageSize = srcSection.PageSetup.PageSize;
            newSection.PageSetup.PageStartingNumber = srcSection.PageSetup.PageStartingNumber;
            newSection.PageSetup.RestartPageNumbering = srcSection.PageSetup.RestartPageNumbering;
            newSection.PageSetup.VerticalAlignment = srcSection.PageSetup.VerticalAlignment;
            //Updates page border.
            newSection.PageSetup.Borders.Bottom.BorderType = srcSection.PageSetup.Borders.Bottom.BorderType;
            newSection.PageSetup.Borders.Bottom.Color = srcSection.PageSetup.Borders.Bottom.Color;
            newSection.PageSetup.Borders.Bottom.LineWidth = srcSection.PageSetup.Borders.Bottom.LineWidth;
            newSection.PageSetup.Borders.Bottom.Shadow = srcSection.PageSetup.Borders.Bottom.Shadow;
            newSection.PageSetup.Borders.Bottom.Space = srcSection.PageSetup.Borders.Bottom.Space;

            newSection.PageSetup.Borders.Top.BorderType = srcSection.PageSetup.Borders.Top.BorderType;
            newSection.PageSetup.Borders.Top.Color = srcSection.PageSetup.Borders.Top.Color;
            newSection.PageSetup.Borders.Top.LineWidth = srcSection.PageSetup.Borders.Top.LineWidth;
            newSection.PageSetup.Borders.Top.Shadow = srcSection.PageSetup.Borders.Top.Shadow;
            newSection.PageSetup.Borders.Top.Space = srcSection.PageSetup.Borders.Top.Space;

            newSection.PageSetup.Borders.Left.BorderType = srcSection.PageSetup.Borders.Left.BorderType;
            newSection.PageSetup.Borders.Left.Color = srcSection.PageSetup.Borders.Left.Color;
            newSection.PageSetup.Borders.Left.LineWidth = srcSection.PageSetup.Borders.Left.LineWidth;
            newSection.PageSetup.Borders.Left.Shadow = srcSection.PageSetup.Borders.Left.Shadow;
            newSection.PageSetup.Borders.Left.Space = srcSection.PageSetup.Borders.Left.Space;

            newSection.PageSetup.Borders.Right.BorderType = srcSection.PageSetup.Borders.Right.BorderType;
            newSection.PageSetup.Borders.Right.Color = srcSection.PageSetup.Borders.Right.Color;
            newSection.PageSetup.Borders.Right.LineWidth = srcSection.PageSetup.Borders.Right.LineWidth;
            newSection.PageSetup.Borders.Right.Shadow = srcSection.PageSetup.Borders.Right.Shadow;
            newSection.PageSetup.Borders.Right.Space = srcSection.PageSetup.Borders.Right.Space;
        }
        /// <summary>
        /// Add new section as sibling of current section
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="entity"></param>
        /// <param name="newSiblings"></param>
        public static void AddSiblings<T>(IEntity entity, IEnumerable<T> newSiblings) where T : class, IEntity
        {
            int newIndex = GetIndex(entity) + 1;

            ICompositeEntity container = entity.Owner as ICompositeEntity;
            if (container == null)
            {
                throw new ApplicationException("Unable to add new siblings to this entity as it does not have a valid container.");
            }

            foreach (var newSibling in newSiblings)
            {
                container.ChildEntities.Insert(newIndex++, newSibling);
            }
        }
        /// <summary>
        /// Get the customers details to perform mail merge.
        /// </summary>
        private static List<ExpandoObject> GetCustomers()
        {
            List<ExpandoObject> customers = new List<ExpandoObject>();
            customers.Add(GetDynamicCustomer(100, "Robert", "Syncfusion"));
            customers.Add(GetDynamicCustomer(102, "John", "Syncfusion"));
            customers.Add(GetDynamicCustomer(110, "David", "Syncfusion"));
            return customers;
        }
        /// <summary>
        /// Get the order details to perform mail merge
        /// </summary>
		private static List<ExpandoObject> GetOrders()
        {
            List<ExpandoObject> orders = new List<ExpandoObject>();
            orders.Add(GetDynamicOrder(1001, "MSWord", 100));
            orders.Add(GetDynamicOrder(1002, "AdobeReader", 100));
            orders.Add(GetDynamicOrder(1003, "VisualStudio", 102));
            return orders;
        }
        /// <summary>
        /// Generate customer details as dynamic objects.
        /// </summary>
        /// <param name="customerID">Represents an customer id</param>
        /// <param name="customerName">Represents a customer name</param>
        /// <param name="companyName">Represents a company name</param>
		private static dynamic GetDynamicCustomer(int customerID, string customerName, string companyName)
        {
            dynamic dynamicCustomer = new ExpandoObject();
            dynamicCustomer.CustomerID = customerID;
            dynamicCustomer.CustomerName = customerName;
            dynamicCustomer.CompanyName = companyName;
            return dynamicCustomer;
        }
        /// <summary>
        /// Generate order details as dynamic objects.
        /// </summary>
        /// <param name="orderID">Represents an order id</param>
        /// <param name="orderName">Represents an order name</param>
        /// <param name="customerID">Represents customer Id</param>
		private static dynamic GetDynamicOrder(int orderID, string orderName, int customerID)
        {
            dynamic dynamicOrder = new ExpandoObject();
            dynamicOrder.OrderID = orderID;
            dynamicOrder.OrderName = orderName;
            dynamicOrder.CustomerID = customerID;
            return dynamicOrder;
        }
        #endregion
    }
}

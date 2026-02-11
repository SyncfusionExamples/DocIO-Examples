using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace MailMerge_Results_in_Two_Columns
{
    class Program
    {
        public static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read))
            {
                using (WordDocument document = new WordDocument(fileStream, FormatType.Docx))
                {
                    // Get student data
                    List<StudentsGroup> studentsGroupList = GetStudentData();
                    //Create mail merge data table
                    MailMergeDataTable dataTable = new MailMergeDataTable("StudentsGroup", studentsGroupList);
                    // Execute nested mail merge
                    document.MailMerge.ExecuteNestedGroup(dataTable);
                    // Split the document into sections based on tables
                    List<WSection> sections = SplitSectionsByTable(document);
                    // Clear existing sections in the document
                    document.Sections.Clear();
                    //Added newly created sections into the document.
                    foreach (WSection section in sections)
                    {
                        document.Sections.Add(section);
                    }
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.                    
                        document.Save(outputFileStream, FormatType.Docx);
                    }                        
                }
            }
        }
        /// <summary>
        /// Splits the given Word document section into multiple sections based on tables.
        /// </summary>
        /// <param name="document">The Word document to split</param>
        /// <returns>A list of sections created from the document.</returns>
        private static List<WSection> SplitSectionsByTable(WordDocument document)
        {
            // Initialize a list to hold the new sections
            List<WSection> sections = new List<WSection>();
            // Iterate through all sections in the document
            foreach (WSection section in document.Sections)
            {
                // Clone the current section.
                WSection clonedSection = section.Clone();
                // Clear child entities from the cloned section. 
                clonedSection.Body.ChildEntities.Clear();
                // Create a new section from the cloned section.
                WSection newSection = clonedSection.Clone();
                // Get the text body of the current section.
                WTextBody textBody = section.Body;
                // Iterate through each child entity in the Text Body
                for (int i = 0; i < textBody.ChildEntities.Count; i++)
                {                    
                    //Accesses the body items (should be either paragraph, table or block content control) as IEntity
                    IEntity bodyItemEntity = textBody.ChildEntities[i];                    
                    // Decides the element type by using EntityType
                    switch (bodyItemEntity.EntityType)
                    {
                        case EntityType.Paragraph:
                            // Clone and add the paragraph to the new section
                            WParagraph paragraph = bodyItemEntity as WParagraph;
                            newSection.Body.ChildEntities.Add(paragraph.Clone());
                            break;
                        case EntityType.Table:
                            // Mark the first row of the table as a header
                            (bodyItemEntity as WTable).Rows[0].IsHeader = true;
                            // Add a paragraph to separate sections
                            newSection.AddParagraph();
                            // Add the current section to the collection
                            sections.Add(newSection);
                            // Create a new section for the table
                            newSection = clonedSection.Clone();                            
                            newSection.BreakCode = SectionBreakCode.NoBreak;
                            // Setup columns (optional)
                            float spacing = 20;
                            float colWidth = newSection.PageSetup.ClientWidth / 2 - spacing;
                            newSection.AddColumn(colWidth, spacing);
                            newSection.Columns[0].Width = colWidth;
                            // Clone and add the table to the new section
                            newSection.Body.ChildEntities.Add(bodyItemEntity.Clone());
                            sections.Add(newSection);
                            // Reset newSection for further processing
                            newSection = clonedSection.Clone();
                            break;
                        case EntityType.BlockContentControl:
                            // Clone and add the block content control to the new section
                            BlockContentControl blockContentControl = bodyItemEntity as BlockContentControl;
                            newSection.Body.ChildEntities.Add((BlockContentControl)blockContentControl.Clone());
                            break;
                    }
                }
            }
            // Return the list of newly created sections
            return sections;
        }
        /// <summary>
        /// Gets the Student details to perform mail merge
        /// </summary>
        /// <returns></returns>
        public static List<StudentsGroup> GetStudentData()
        {
            List<Student> students = new List<Student>();
            List<Student> students1 = new List<Student>();
            List<Student> students2 = new List<Student>();
            for (int i = 1; i <= 45; i++)
            {
                students.Add(new Student
                {
                    RollNo = $"{i}",
                    AdmissionNo = $"ADM{i:000}",
                    StudentName = $"Class 1A Student {i}",
                    Marks = $"M{i:000}",
                });
            }
            for (int i = 1; i <= 45; i++)
            {
                students1.Add(new Student
                {
                    RollNo = $"{i}",
                    AdmissionNo = $"ADM{i:000}",
                    StudentName = $"Class 2A Student {i}",
                    Marks = $"M{i:000}",
                });
            }
            for (int i = 1; i <= 45; i++)
            {
                students2.Add(new Student
                {
                    RollNo = $"{i}",
                    AdmissionNo = $"ADM{i:000}",
                    StudentName = $"Class 2B Student {i}",
                    Marks = $"M{i:000}",
                });
            }
            List<StudentsGroup> parentList = new List<StudentsGroup>
            {
                new StudentsGroup {
                    Class="1#A",
                    Exam="MidTerm",
                    Students = students
                },
                new StudentsGroup {
                    Class="2#A",
                    Exam="MidTerm",
                    Students = students1
                },
                  new StudentsGroup {
                    Class="2#B",
                    Exam="MidTerm",
                    Students = students2
                }
            };

            return parentList;
        }
    }
    /// <summary>
    /// Represents a class to maintain Student details
    /// </summary>
    public class Student
    {
        public string RollNo { get; set; }
        public string AdmissionNo { get; set; }
        public string StudentName { get; set; }
        public string Marks { get; set; }
    }
    public class StudentsGroup
    {
        public string Class { get; set; }
        public string Exam { get; set; }
        public List<Student> Students { get; set; }
    }
}
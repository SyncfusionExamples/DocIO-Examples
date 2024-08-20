using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.IO;

namespace Get_section_number_of_bookmark
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Input.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Open an existing Word document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic))
                {
                    // Get the bookmark instance by using FindByName method of BookmarkCollection with bookmark name
                    Bookmark bookmark = document.Bookmarks.FindByName("BOOKMARK2");
                    if (bookmark != null)
                    {
                        // Get the owner section of bookmark.
                        WSection section = GetOwnerEntity(bookmark.BookmarkStart) as WSection;
                        if (section != null)
                        {
                            // Get the index value of section.
                            int sectionIndex = document.ChildEntities.IndexOf(section);
                            int sectionNumber = sectionIndex + 1;
                            Console.WriteLine("This bookmark would be in section " + sectionNumber);
                            Console.ReadKey();
                        }
                    }
                }
                
            }
        }
        /// <summary>
        /// Get the Entity owner
        /// </summary>
        private static Entity GetOwnerEntity(BookmarkStart bookmarkStart)
        {
            Entity baseEntity = bookmarkStart.Owner;

            while (!(baseEntity is WSection))
            {
                if (baseEntity is null)
                    return baseEntity;
                baseEntity = baseEntity.Owner;
            }
            return baseEntity;
        }
    }
}

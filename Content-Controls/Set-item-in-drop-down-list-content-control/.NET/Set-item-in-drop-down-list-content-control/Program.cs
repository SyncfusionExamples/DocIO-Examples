using Syncfusion.DocIO.DLS;

namespace Set_item_in_drop_down_list_content_control
{
    class Program
    {
        static void Main(string []args)
        {
            WordDocument document = new WordDocument(Path.GetFullPath("../../../Data/Dropdown.docx"));
            List<Entity> contentControls = document.FindAllItemsByProperty(EntityType.InlineContentControl, null, null);
            List<Entity> blockContentControls = document.FindAllItemsByProperty(EntityType.BlockContentControl, null, null);

            if(contentControls != null)
            {
                foreach(Entity icc in contentControls)
                {
                    InlineContentControl contentControl = icc as InlineContentControl;
                    if (contentControl.ContentControlProperties.Type == ContentControlType.DropDownList)
                        SelectDropDownOption(contentControl, 1);
                }
            }
            if (blockContentControls != null)
            {
                foreach (Entity bcc in blockContentControls)
                {
                    BlockContentControl blockContentControl = bcc as BlockContentControl;
                    if (blockContentControl.ContentControlProperties.Type == ContentControlType.DropDownList)
                        SelectDropDownOption(blockContentControl, 1);
                }
            }
            //Saves and closes the document instance
            document.Save(Path.GetFullPath("../../../Output/output.docx"));
            document.Close();
        }

        /// <summary>
        /// Selects the specified item in an inline dropdown list content control
        /// and updates the displayed text with the selected item's display value.
        /// </summary>
        /// <param name="dropdown">The inline dropdown list content control.</param>
        /// <param name="index">The index of the dropdown item to select.</param>
        private static void SelectDropDownOption(InlineContentControl dropdown, int index)
        {
            if (dropdown.ContentControlProperties.Type == ContentControlType.DropDownList)
            {
                WTextRange selectedText = dropdown.ParagraphItems[0] as WTextRange;
                selectedText.Text = dropdown.ContentControlProperties.ContentControlListItems[index].DisplayText;
                selectedText.CharacterFormat.TextColor = Syncfusion.Drawing.Color.Black;
            }
        }

        /// <summary>
        /// Selects the specified item in a block dropdown list content control
        /// and updates the displayed text with the selected item's display value.
        /// </summary>
        /// <param name="dropdown">The block dropdown list content control.</param>
        /// <param name="index">The index of the dropdown item to select.</param>
        private static void SelectDropDownOption(BlockContentControl dropdown, int index)
        {
            if (dropdown.ContentControlProperties.Type == ContentControlType.DropDownList)
            {
                WTextRange selectedText = dropdown.TextBody.Paragraphs[0].ChildEntities[0] as WTextRange;
                selectedText.Text = dropdown.ContentControlProperties.ContentControlListItems[index].DisplayText;
                selectedText.CharacterFormat.TextColor = Syncfusion.Drawing.Color.Black;
            }
        }
    }
}


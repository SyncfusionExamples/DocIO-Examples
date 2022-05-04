#region Copyright Syncfusion Inc. 2001 - 2017
//
//  Copyright Syncfusion Inc. 2001 - 2017. All rights reserved.
//
//  Use of this code is subject to the terms of our license.
//  A copy of the current license can be obtained at any time by e-mailing
//  licensing@syncfusion.com. Any infringement will be prosecuted under
//  applicable laws. 
//
#endregion
using System;
using System.Windows;
using System.Windows.Media;
using Syncfusion.DocIO.DLS;
using System.IO;
using Syncfusion.DocIO;
using System.ComponentModel;
using Microsoft.Win32;
using Syncfusion.Windows.Shared;
using Syncfusion.DocToPDFConverter;
using Syncfusion.Pdf;
using System.Collections.Generic;
using System.Windows.Forms;

namespace GenerateRandomMathsQuestions
{
    /// <summary>
    /// Interaction logic for Window1.xaml
    /// </summary>
    public partial class Window1 : ChromelessWindow
    {
        Microsoft.Win32.OpenFileDialog openFileDialog1 = new Microsoft.Win32.OpenFileDialog();
        static string file1, file2, path;
        # region Constructor
        /// <summary>
        /// Window constructor
        /// </summary>
        public Window1()
        {
            InitializeComponent();
            ImageSourceConverter img = new ImageSourceConverter();
            image1.Source = (ImageSource)img.ConvertFromString(@"..\..\DocIO\docio_header.png");
            this.Icon = (ImageSource)img.ConvertFromString(@"..\..\DocIO\sfLogo.ico");
            path = @"..\..\DocIO\";
            textBox1.Text = "Choose File";      
            textBox2.Text = "Choose Output path";
            openFileDialog1.InitialDirectory = new DirectoryInfo(path).FullName;
            openFileDialog1.Filter = "Word Document(*.doc *.docx *.rtf)|*.doc;*.docx;*.rtf";
       
        }
        # endregion

        # region Events
        /// <summary>
        /// Creates word document
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                WordDocument document = new WordDocument(this.textBox1.Text);
                WSection section = document.Sections[0];
                WTable table = section.Tables[0] as WTable;
                //Gets the random numbers
                Random questionNum = new Random();
                List<int> listNumbers = new List<int>();
                int number;
                //Console.WriteLine("Enter the random number count:");
                int randomNumber = Convert.ToInt32(this.textBox3.Text);
                for (int i = 0; i <= randomNumber; i++)
                {
                    do
                    {
                        number = questionNum.Next(1, 50);
                    } while (listNumbers.Contains(number));
                    listNumbers.Add(number);
                }
                string date = DateTime.Now.ToString("dd-mm-yyyy HH-mm-ss");
                //Gets the document with random questions
                GenerateQuestionDocument(table, listNumbers, this.textBox2.Text + @"\Questions-"+ date + ".docx");
                //Gets the document with answers
                GenerateAnswerDocument(table, listNumbers, this.textBox2.Text + @"\Answers-" + date + ".docx");
                document.Close();
                System.Windows.MessageBox.Show("The questions and answer key documents have been generated in the following location: " + "["+ this.textBox2.Text + "]");
               
            }
            catch (Exception Ex)
            {
                System.Windows.MessageBox.Show(Ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
      
        /// <summary>
        /// Handles the Click event of the btnBrowse control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.Windows.RoutedEventArgs"/> instance containing the event data.</param>
        private void btnBrowse_Click(object sender, RoutedEventArgs e)
        {
            openFileDialog1.FileName = "";

            if (openFileDialog1.ShowDialog().Value)
            {
                this.textBox1.Text = openFileDialog1.FileName;
                file1 = openFileDialog1.FileName;
            }
        }
		
        /// <summary>
        /// Handles the Click event of the btnBrowse1 control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.Windows.RoutedEventArgs"/> instance containing the event data.</param>
        private void btnBrowse1_Click(object sender, RoutedEventArgs e)
        {
            openFileDialog1.FileName = "";
            var fbd = new FolderBrowserDialog();
            DialogResult result = fbd.ShowDialog();
            if (result.ToString() == "OK" )
            {
                this.textBox2.Text = fbd.SelectedPath;
                file2 = openFileDialog1.InitialDirectory;
            }         
        }
        #endregion


        /// <summary>
        /// Get the document with questions.
        /// </summary>
        private static void GenerateAnswerDocument(WTable wTable, List<int> num, String saveFilepath)
        {
            //Creates the document
            WordDocument tempDocument = new WordDocument(path + @"\Answers Template.docx");
            CreateTable(num.Count, tempDocument);
            WTable table = tempDocument.Sections[0].Tables[0] as WTable;
            table.AutoFit(AutoFitType.FitToContent);
            WParagraph paragraph = table.Rows[0].Cells[1].AddParagraph() as WParagraph;
            paragraph.ParagraphFormat.BeforeSpacing = 8f;
            paragraph.ParagraphFormat.AfterSpacing = 8f;
            //Adds text to the cell
            WTextRange textRange = paragraph.AppendText("Answers") as WTextRange;
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 16;
            textRange.CharacterFormat.FontName = "Calibri";
            //Add third cell.
            paragraph = table.Rows[0].AddCell().AddParagraph() as WParagraph;
            paragraph.ParagraphFormat.BeforeSpacing = 8f;
            paragraph.ParagraphFormat.AfterSpacing = 8f;
            //Adds text to the cell
            textRange = paragraph.AppendText("QB Id") as WTextRange;
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 16;
            textRange.CharacterFormat.FontName = "Calibri";

            for (int j = 1; j < num.Count; j++)
            {
                table.Rows[j].AddCell();
                table.Rows[j].RowFormat.IsBreakAcrossPages = false;
                table.Rows[j].Cells[0].AddParagraph().AppendText(j.ToString());
                //Gets the random number
                int randomNum = num[j - 1];
                foreach (WParagraph required in wTable.Rows[randomNum].Cells[2].Paragraphs)
                {
                    //Insert the answers to the cells
                    table.Rows[j].Cells[1].Paragraphs.Add(required.Clone() as WParagraph);
                }
                foreach (WParagraph required in wTable.Rows[randomNum].Cells[0].Paragraphs)
                {
                    //Insert the id to the cells
                    table.Rows[j].Cells[2].Paragraphs.Add(required.Clone() as WParagraph);
                }
            }
            //Saves and close the Word document
            tempDocument.Save(saveFilepath);
            tempDocument.Close();
        }
		
        /// <summary>
        /// Get the document with answer.
        /// </summary>
        private static void GenerateQuestionDocument(WTable wTable, List<int> num, String saveFilepath)
        {

            //Creates the document
            WordDocument tempDocument = new WordDocument(path + @"\Questions Template.docx");
            //Adds a section into Word document
            IWSection section = tempDocument.LastSection;
            section.AddParagraph();
            //Adds a new table into Word document
            IWTable table = section.AddTable();
            //Specifies the total number of rows & columns
            table.ResetCells(num.Count, 2);
            //Apply cell width.
            table.Rows[0].Cells[0].Width = 24;
            table.Rows[0].Cells[1].Width = 444.8f;
            table.TableFormat.Borders.BorderType = Syncfusion.DocIO.DLS.BorderStyle.None; 
           
            for (int j = 1; j < num.Count; j++)
            {
                
                table.Rows[j].RowFormat.IsBreakAcrossPages = false;
                table.Rows[j].Cells[0].AddParagraph().AppendText(j.ToString());
                //Apply cell width.
                table.Rows[j].Cells[0].Width = 24f;
                //Apply cell width.
                table.Rows[j].Cells[1].Width = 444.8f;
                //Gets the random number
                int randomNum = num[j - 1];
                string text = wTable.Rows[randomNum].Cells[0].Paragraphs[0].Items[0].ToString();
                if (wTable.Rows[randomNum].Cells[1].ChildEntities[0] is WTable)
                {
                    foreach (WTable required in wTable.Rows[randomNum].Cells[1].Tables)
                    {
                        //Insert the question to the cells
                        table.Rows[j].Cells[1].Tables.Add(required.Clone());                
                    }
                }
                else if (wTable.Rows[randomNum].Cells[1].ChildEntities[0] is WParagraph)
                {
                    foreach (WParagraph required in wTable.Rows[randomNum].Cells[1].Paragraphs)
                    {
                        //Insert the question to the cells
                        table.Rows[j].Cells[1].Paragraphs.Add(required.Clone() as WParagraph);
                    }
                }
                table.Rows[j].Cells[1].LastParagraph.ParagraphFormat.AfterSpacing = 6f;
            }
            //Saves and close the Word document
            tempDocument.Save(saveFilepath);
            tempDocument.Close();
        }

        /// <summary>
        /// Create the table
        /// </summary>
        /// <param name="num"></param>
        /// <returns></returns>
        private static WordDocument CreateTable(int num, WordDocument document)
        {
            //Adds a section into Word document
            IWSection section = document.LastSection;
            section.AddParagraph();
            //Adds a new table into Word document
            IWTable table = section.AddTable();
            //Specifies the total number of rows & columns
            table.ResetCells(num, 2);
            table.Rows[0].IsHeader = true;
            //Adds text to the first cell
            WParagraph paragraph = table.Rows[0].Cells[0].AddParagraph() as WParagraph;
            paragraph.ParagraphFormat.BeforeSpacing = 8f;
            paragraph.ParagraphFormat.AfterSpacing = 8f;
            WTextRange textRange = paragraph.AppendText("S.No") as WTextRange;
            textRange.CharacterFormat.Bold = true;
            textRange.CharacterFormat.FontSize = 16;
            textRange.CharacterFormat.FontName = "Calibri";
            return document;
        }
    }
}
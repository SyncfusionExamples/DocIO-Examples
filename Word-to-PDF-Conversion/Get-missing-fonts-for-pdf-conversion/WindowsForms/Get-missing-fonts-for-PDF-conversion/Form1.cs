using Syncfusion.DocIO.DLS;
using Syncfusion.DocToPDFConverter;
using Syncfusion.OfficeChartToImageConverter;
using Syncfusion.Pdf;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Get_missing_fonts_for_PDF_conversion
{
    public partial class Form1 : Form
    {
        // List to store names of fonts that are not installed
        static List<string> fonts = new List<string>();
        public Form1()
        {
            InitializeComponent();
        }

        private void browseButton_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = Application.StartupPath + @"..\..\Data\DocIO\";
            openFileDialog1.FileName = "";
            DialogResult result = openFileDialog1.ShowDialog();

            if (result == DialogResult.OK)
            {
                this.textBox1.Text = openFileDialog1.SafeFileName;
                this.textBox1.Tag = openFileDialog1.FileName;
            }
        }

        private void btnconvert_Click_1(object sender, EventArgs e)
        {
            if (this.textBox1.Text != String.Empty)
            {
                WordDocument wordDoc = new WordDocument((string)textBox1.Tag, Syncfusion.DocIO.FormatType.Automatic);
                //Initialize chart to image converter for converting charts in word to pdf conversion
                wordDoc.ChartToImageConverter = new ChartToImageConverter();
                wordDoc.ChartToImageConverter.ScalingMode = Syncfusion.OfficeChart.ScalingMode.Normal;
                DocToPDFConverter converter = new DocToPDFConverter();

                // Hook the font substitution event to detect missing fonts
                wordDoc.FontSettings.SubstituteFont += FontSettings_SubstituteFont;
                //Convert word document into PDF document
                PdfDocument pdfDoc = converter.ConvertToPDF(wordDoc);
                //Save the pdf file
                pdfDoc.Save("DoctoPDF.pdf");

                // Print the fonts that are not available in machine, but used in Word document.
                if (fonts.Count > 0)
                {
                    string misddedFonts = string.Empty;
                    Console.WriteLine("Fonts not available in environment:");
                    int i = 0;
                    foreach (string font in fonts)
                    {
                        misddedFonts += '\n' + font;
                        i++;

                    }
                    MessageBox.Show("Fonts not available in environment:" + misddedFonts);
                }
                else
                {
                    MessageBox.Show("Fonts used in Word document are available in environment.");
                }
                //Message box confirmation to view the created document.
                if (MessageBox.Show("Do you want to view the generated PDF?", " Document has been created", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                {
                    try
                    {                     
                        System.Diagnostics.Process.Start("DoctoPDF.pdf");
                        //Exit
                        this.Close();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.ToString());
                    }
                }
            }
            else
            {
                MessageBox.Show("Browse a word document and click the button to convert as a PDF.");
            }
        }

        // Event handler for font substitution event
        static void FontSettings_SubstituteFont(object sender, SubstituteFontEventArgs args)
        {
            // Add the original font name to the list if it's not already there
            if (!fonts.Contains(args.OriginalFontName))
                fonts.Add(args.OriginalFontName);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.textBox1.Text = "Input.docx";
            this.textBox1.Tag = Application.StartupPath + @"..\..\..\Data\Input.docx";
        }
    }
}

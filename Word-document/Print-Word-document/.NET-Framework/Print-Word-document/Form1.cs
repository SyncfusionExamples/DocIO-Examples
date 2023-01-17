using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.IO;
using System.Diagnostics;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Windows.Forms;
using Syncfusion.Windows.Forms.Tools;
using System.Drawing.Printing;

namespace EssentialDocIOSamples
{

    public class Form1 : MetroForm
    {
        #region Private Members
        private System.Windows.Forms.OpenFileDialog openFileDialog;
        private Syncfusion.Windows.Forms.ButtonAdv browseButton;
        private System.Windows.Forms.TextBox textBox;
        private System.Windows.Forms.Label browseLabel;
        private System.Windows.Forms.PictureBox pictureBox;
        private Syncfusion.Windows.Forms.ButtonAdv printButton;
        private System.Windows.Forms.Label descriptionLabel;
        private IContainer components;
        System.Drawing.Image[] images = null;
        int startPageIndex = 0;
        int endPageIndex = 0;
        string dataPath;
        #endregion

        #region Constructor, Main and Dispose
        public Form1()
        {
            //
            // Required for Windows Form Designer support
            //
            InitializeComponent();
            Application.EnableVisualStyles();

            //
            // TODO: Add any constructor code after InitializeComponent call
            //
        }


        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.Run(new Form1());
        }
        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (components != null)
                {
                    components.Dispose();
                }
            }
            base.Dispose(disposing);
        }
        #endregion

        #region Windows Form Designer generated code
        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.browseButton = new Syncfusion.Windows.Forms.ButtonAdv();
            this.textBox = new System.Windows.Forms.TextBox();
            this.browseLabel = new System.Windows.Forms.Label();
            this.pictureBox = new System.Windows.Forms.PictureBox();
            this.printButton = new Syncfusion.Windows.Forms.ButtonAdv();
            this.descriptionLabel = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox)).BeginInit();
            this.SuspendLayout();
            // 
            // openFileDialog
            // 
            this.openFileDialog.FileName = "openFileDialog1";
            // 
            // browseButton
            // 
            this.browseButton.BeforeTouchSize = new System.Drawing.Size(21, 22);
            this.browseButton.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.browseButton.IsBackStageButton = false;
            this.browseButton.Location = new System.Drawing.Point(332, 157);
            this.browseButton.Name = "browseButton";
            this.browseButton.Size = new System.Drawing.Size(21, 22);
            this.browseButton.TabIndex = 92;
            this.browseButton.Text = ". . .";
            this.browseButton.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.browseButton.UseVisualStyleBackColor = true;
            this.browseButton.Click += new System.EventHandler(this.button3_Click);
            // 
            // textBox
            // 
            this.textBox.Location = new System.Drawing.Point(6, 157);
            this.textBox.Name = "textBox";
            this.textBox.Size = new System.Drawing.Size(307, 22);
            this.textBox.TabIndex = 91;
            // 
            // browseLabel
            // 
            this.browseLabel.AutoSize = true;
            this.browseLabel.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(161)));
            this.browseLabel.Location = new System.Drawing.Point(3, 135);
            this.browseLabel.Name = "browseLabel";
            this.browseLabel.Size = new System.Drawing.Size(181, 13);
            this.browseLabel.TabIndex = 90;
            this.browseLabel.Text = "Browse a Word Document :";
            // 
            // pictureBox
            // 
            this.pictureBox.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox.Image")));
            this.pictureBox.Location = new System.Drawing.Point(0, 0);
            this.pictureBox.Name = "pictureBox";
            this.pictureBox.Size = new System.Drawing.Size(365, 82);
            this.pictureBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox.TabIndex = 89;
            this.pictureBox.TabStop = false;
            // 
            // printButton
            // 
            this.printButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.printButton.Appearance = Syncfusion.Windows.Forms.ButtonAppearance.Metro;
            this.printButton.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(17)))), ((int)(((byte)(158)))), ((int)(((byte)(218)))));
            this.printButton.BeforeTouchSize = new System.Drawing.Size(78, 26);
            this.printButton.BorderStyleAdv = Syncfusion.Windows.Forms.ButtonAdvBorderStyle.Dashed;
            this.printButton.ComboEditBackColor = System.Drawing.Color.Silver;
            this.printButton.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.printButton.ForeColor = System.Drawing.Color.White;
            this.printButton.IsBackStageButton = false;
            this.printButton.KeepFocusRectangle = false;
            this.printButton.Location = new System.Drawing.Point(275, 193);
            this.printButton.Name = "printButton";
            this.printButton.Office2007ColorScheme = Syncfusion.Windows.Forms.Office2007Theme.Managed;
            this.printButton.Size = new System.Drawing.Size(78, 26);
            this.printButton.TabIndex = 13;
            this.printButton.Text = "Print";
            this.printButton.UseVisualStyle = true;
            this.printButton.UseVisualStyleBackColor = false;
            this.printButton.Click += new System.EventHandler(this.button1_Click);
            // 
            // descriptionLabel
            // 
            this.descriptionLabel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.descriptionLabel.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.descriptionLabel.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(64)))));
            this.descriptionLabel.Location = new System.Drawing.Point(0, 85);
            this.descriptionLabel.Name = "descriptionLabel";
            this.descriptionLabel.Size = new System.Drawing.Size(363, 51);
            this.descriptionLabel.TabIndex = 97;
            this.descriptionLabel.Text = "Click the below button to print the word document.In this,Essential DocIO render " +
    "the word document contents page by page as images and print the rendered image " +
    "using PrintDialog.";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(17)))), ((int)(((byte)(158)))), ((int)(((byte)(218)))));
            this.ClientSize = new System.Drawing.Size(365, 225);
            this.Controls.Add(this.descriptionLabel);
            this.Controls.Add(this.printButton);
            this.Controls.Add(this.browseButton);
            this.Controls.Add(this.textBox);
            this.Controls.Add(this.browseLabel);
            this.Controls.Add(this.pictureBox);
            this.DropShadow = true;
            this.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Print";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        #endregion

        # region Form Load
        /// <summary>
        /// Handles the Load event of the Form1 control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
        private void Form1_Load(object sender, EventArgs e)
        {
            this.textBox.Text = "DocToImage.docx";
            dataPath = new DirectoryInfo(Application.StartupPath + @"../../../Data/").FullName;
            this.textBox.Tag = dataPath + "DocToImage.docx";
        }
        #endregion

        #region Browse a Word document
        /// <summary>
        /// Browse the word document to print
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {
            openFileDialog.InitialDirectory = dataPath;
            openFileDialog.FileName = "";
            DialogResult result = openFileDialog.ShowDialog();
            if (result == DialogResult.OK)
            {
                this.textBox.Text = openFileDialog.SafeFileName;
                this.textBox.Tag = openFileDialog.FileName;
            }
        }
        #endregion
        
        #region Print Button click event
        /// <summary>
        /// Print the word document
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, System.EventArgs e)
        {
            //Loads the image for showing document rendering progress.
            this.descriptionLabel.Image = Image.FromFile(new DirectoryInfo(Application.StartupPath + @"../../../Images/Animation.gif").FullName);
            BackgroundWorker worker = new BackgroundWorker();
            //Hooks DoWork event.
            worker.DoWork += worker_DoWork;
            //Hooks run worker completed event.
            worker.RunWorkerCompleted += worker_RunWorkerCompleted;
            //Runs the background worker.
            worker.RunWorkerAsync();
        }
        #endregion

        #region PrintPage event
        /// <summary>
        /// Handles the OnPrintPage event to draw document pages.
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The <see cref="PrintPageEventArgs"/> instance containing the event data.</param>
        private void OnPrintPage(object sender, PrintPageEventArgs e)
        {
            //Gets the print start page width.
            int currentPageWidth = images[startPageIndex].Width;
            //Gets the print start page height.
            int currentPageHeight = images[startPageIndex].Height;
            //Gets the visible bounds width for print.
            int visibleClipBoundsWidth = (int)e.Graphics.VisibleClipBounds.Width;
            //Gets the visible bounds height for print.
            int visibleClipBoundsHeight = (int)e.Graphics.VisibleClipBounds.Height;
            //Checks if the page layout is landscape or portrait.
            if (currentPageWidth > currentPageHeight)
            {
                //Translates the position.
                e.Graphics.TranslateTransform(0, visibleClipBoundsHeight);
                //Rotates the object at 270 degrees
                e.Graphics.RotateTransform(270.0f);
                //Draws the current page image.
                e.Graphics.DrawImage(images[startPageIndex], new System.Drawing.Rectangle(0, 0, currentPageWidth, currentPageHeight));
            }
            else
            {
                //Draws the current page image.
                e.Graphics.DrawImage(images[startPageIndex], new System.Drawing.Rectangle(0, 0, visibleClipBoundsWidth, visibleClipBoundsHeight));
            }
            //Disposes the current page image after drawing.
            images[startPageIndex].Dispose();
            //Increments the start page index.
            startPageIndex++;
            //Updates if the document contains some more pages to print.
            if (startPageIndex < endPageIndex)
                e.HasMorePages = true;
            else
                startPageIndex = 0;
        }
        # endregion

        #region Print document
        /// <summary>
        /// Handles the RunWorkerCompleted event to print the document.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="RunWorkerCompletedEventArgs" /> instance containing the event data.</param>
        private void worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (this.textBox.Text != String.Empty)
            {
                //Creates new PrintDialog instance.
                System.Windows.Forms.PrintDialog printDialog = new System.Windows.Forms.PrintDialog();
                //Sets new PrintDocument instance to print dialog.
                printDialog.Document = new PrintDocument();
                //Enables the print current page option.
                printDialog.AllowCurrentPage = true;
                //Enables the print selected pages option.
                printDialog.AllowSomePages = true;
                //Sets the start and end page index
                printDialog.PrinterSettings.FromPage = 1;
                printDialog.PrinterSettings.ToPage = images.Length;
                //Opens the print dialog box.
                if (printDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    //Checks if the selected page range is valid.
                    if (printDialog.PrinterSettings.FromPage > 0 && printDialog.PrinterSettings.ToPage <= images.Length)
                    {
                        //Updates the start page of the document to print.
                        startPageIndex = printDialog.PrinterSettings.FromPage - 1;
                        //Updates the end page of the document to print.
                        endPageIndex = printDialog.PrinterSettings.ToPage;
                        //Hooks the PrintPage event to handle be drawing pages for printing.
                        printDialog.Document.PrintPage += new PrintPageEventHandler(OnPrintPage);
                        //Prints the document.
                        printDialog.Document.Print();
                    }
                    else
                    {
                        MessageBoxAdv.Show("The page range is invalid" + Environment.NewLine + "Enter numbers between 1 and " + images.Length.ToString(), "Print Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    //Disposes the print dialog.s
                    printDialog.Dispose();
                    //Exits the form.
                    this.Close();
                }
            }
            else
            {
                MessageBoxAdv.Show("Browse a word document and click the button to Print the Word document.");
            }
            if (sender is BackgroundWorker)
                //Unhooks run worker completed event.
                (sender as BackgroundWorker).RunWorkerCompleted -= worker_RunWorkerCompleted;
        }
        #endregion

        # region Render document
        /// <summary>
        /// Handles the DoWork event to render the Word document as image.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="DoWorkEventArgs"/> instance containing the event data.</param>
        private void worker_DoWork(object sender, DoWorkEventArgs e)
        {
            //Opens the Word document.
            WordDocument wordDoc = new WordDocument((string)this.textBox.Tag);
            //Renders the Word document as image.
            images = wordDoc.RenderAsImages(ImageType.Metafile);
            endPageIndex = images.Length;
            //Closes the Word Document.
            wordDoc.Close();
            //Disposes the label image.
            if (this.descriptionLabel.Image != null)
                this.descriptionLabel.Image.Dispose();
            this.descriptionLabel.Image = null;
            if (sender is BackgroundWorker)
                //Unhooks do work event.
                (sender as BackgroundWorker).DoWork -= worker_DoWork;
        }
        #endregion
    }
}

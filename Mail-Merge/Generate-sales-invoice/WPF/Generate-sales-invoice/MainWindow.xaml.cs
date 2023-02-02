using Syncfusion.DocIO.DLS;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Windows;
using System.Windows.Media;

namespace Generate_sales_invoice
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //Sets the name or network address of the instance of SQL Server to connect
        string datasource = @"DataSource";
        //Sets the user ID to be used when connecting to SQL Server.
        string userID = "Your User ID";
        //Sets the password for the SQL Server account.
        string password = "Your Password";

        SqlConnection connection = null;
        List<int> orderIDs = null;

        public MainWindow()
        {
            //Initializes the component
            InitializeComponent();
            //Creates SqlConnection to connect with database
            CreateSqlConnection();
            //Read all the OrderIDs from database.
            ReadOrderIDs();

            //Displays list of OrderIDs in combo box.
            foreach (int id in orderIDs)
                ComboBox1.Items.Add(id.ToString());
            ComboBox1.SelectedIndex = 0;

            //Sets image and icon for application.
            ImageSourceConverter img = new ImageSourceConverter();
            image1.Source = (ImageSource)img.ConvertFromString(@"..\..\Images\docio_header.png");
            this.Icon = (ImageSource)img.ConvertFromString(@"..\..\Images\sfLogo.ico");
        }

        #region Button click event
        /// <summary>
        /// Generates sales invoice
        /// </summary>
        private void Generate_Invoice_click(object sender, RoutedEventArgs e)
        {
            //Get the input file path.
            string dataPath = @"..\..\Data\SalesInvoiceDemo.docx";
            //Gets the selected order Id.
            int selectedID = orderIDs[ComboBox1.SelectedIndex];

            //Opens an existing template Word document
            WordDocument document = new WordDocument(dataPath);
            //Get commands to execute in database
            ArrayList commands = GetCommands(selectedID);
            //Executes nested mail merge.
            document.MailMerge.ExecuteNestedGroup(connection, commands, true);
            //Disposes the database connection
            connection.Dispose();
            //Saves and closes the Word document.
            document.Save(@"..\..\Sample.docx");
            document.Close();

            //Clears the collection.
            orderIDs.Clear();
            //Exit the application
            Close();
        }
        #endregion

        #region Helper Methods
        /// <summary>
        /// Create commands to execute in database during mail merge process.
        /// </summary>
        /// <param name="OrderID">Represents an OrderID to generate sales invoice. </param>
        /// <returns></returns>
        private ArrayList GetCommands(int OrderID)
        {
            //Creates a collection.
            ArrayList commands = new ArrayList();
            //Creates commands and add in collection.
            DictionaryEntry entry = new DictionaryEntry("Orders", string.Format("SELECT * FROM Orders WHERE OrderID={0}", OrderID.ToString()));
            commands.Add(entry);
            entry = new DictionaryEntry("OrderDetails", string.Format("SELECT * FROM OrderDetails WHERE OrderID={0}", OrderID.ToString()));
            commands.Add(entry);
            entry = new DictionaryEntry("OrderTotals", string.Format("SELECT * FROM OrderTotals WHERE OrderID={0}", OrderID.ToString()));
            commands.Add(entry);
            return commands;
        }
        /// <summary>
        /// Read OrderID column in database table
        /// </summary>
        /// <returns>Returns list of order IDs</returns>
        private List<int> ReadOrderIDs()
        {
            orderIDs = new List<int>();
            if (connection.State == System.Data.ConnectionState.Open)
            {
                //Sql query to retrieve all OrderID from table
                string query = "SELECT OrderID FROM Orders";
                //Creates a SqlCommand to execute in database.
                SqlCommand command = new SqlCommand(query, connection);
                //Read all order Ids from database.
                SqlDataReader reader = command.ExecuteReader();
                //Read each records.
                while (reader.Read())
                    orderIDs.Add(reader.GetInt32(0));
                //Closes the reader.
                reader.Close();
            }
            return orderIDs;
        }
        /// <summary>
        /// Creates connection to SQL Server database
        /// </summary>
        private void CreateSqlConnection()
        {
            try
            {
                // Build connection string
                SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();
                builder.DataSource = datasource;
                builder.UserID = userID;
                builder.Password = password;
                //Sets database file path.
                builder.AttachDBFilename = System.IO.Path.GetFullPath(@"..\..\Data\InvoiceDetails.mdf");
                builder.PersistSecurityInfo = true;
                builder.IntegratedSecurity = true;                
                connection = new SqlConnection(builder.ConnectionString);
                //Opens the Sql connection
                connection.Open();
            }
            catch (Exception e)
            {
                //Message box to show error while creating SqlConnection.
                if (MessageBox.Show("Please enter valid data source and credentials to create SqlConnection.", "Database Connection Error",
                        MessageBoxButton.OK, MessageBoxImage.Information) == MessageBoxResult.OK)
                {
                    //Exit
                    Close();
                }
            }
        }
        #endregion
    }
}

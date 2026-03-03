using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Generate_loan_agreement
{
    class Program
    {
        static void Main(string[] args)
        {
            // Generate loan agreements with combined all records.
            GenerateAgreementsDocuments();
            // Generate individual agreement documents.
            GenerateLoanAgreements();
        }
        /// <summary>
        /// Generates loan agreement documents using mail merge from template.
        /// Creates combined document with all records.
        /// </summary>
        public static void GenerateAgreementsDocuments()
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.ReadWrite))
            {
                //Loads an existing Word document into DocIO instance.
                using (WordDocument document = new WordDocument(fileStream, FormatType.Automatic))
                {
                    //Gets the employee details as IEnumerable collection.
                    List<AgreementDetails> agreementList = GetAgreementDeatils();
                    //Creates an instance of MailMergeDataTable by specifying MailMerge group name and IEnumerable collection.
                    MailMergeDataTable dataSource = new MailMergeDataTable("AgreementDetails", agreementList);
                    //Enable the boolean to remove empty paragraph and start each record in new page.
                    document.MailMerge.RemoveEmptyParagraphs = true;
                    document.MailMerge.StartAtNewPage = true;
                    //Performs Mail merge.
                    document.MailMerge.ExecuteGroup(dataSource);
                    //Creates file stream.
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"../../../Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            }
        }
        /// <summary>
        /// Generates individual loan agreement documents, one document per record.
        /// </summary>
        public static void GenerateLoanAgreements()
        {
            string templatePath = Path.GetFullPath(@"Data/Template.docx");
            string outputDir = Path.GetFullPath(@"../../../Output");
            // Get all agreement details.
            List<AgreementDetails> agreementList = GetAgreementDeatils();
            //Iterate each record and generate document.
            foreach (var agreement in agreementList)
            {
                using (FileStream fileStream = new FileStream(templatePath, FileMode.Open, FileAccess.Read))
                {
                    //Loads an existing Word document into DocIO instance.
                    using (WordDocument document = new WordDocument(fileStream, FormatType.Automatic))
                    {
                        // Create a list with single record
                        List<AgreementDetails> singleRecord = new List<AgreementDetails> { agreement };
                        // Create mail merge data source with single record
                        MailMergeDataTable dataSource = new MailMergeDataTable("AgreementDetails", singleRecord);
                        //Enable the boolean to remove empty paragraph.
                        document.MailMerge.RemoveEmptyParagraphs = true;
                        // Execute mail merge for this single record
                        document.MailMerge.ExecuteGroup(dataSource);
                        // Create unique filename using AgreementNumber
                        string fileName = $"LoanAgreement_{agreement.AgreementNumber}.docx";
                        string outputPath = Path.Combine(outputDir, fileName);
                        //Creates file stream.
                        using (FileStream outputStream = new FileStream(outputPath, FileMode.Create, FileAccess.ReadWrite))
                        {
                            //Saves the Word document to file stream.
                            document.Save(outputStream, FormatType.Docx);
                        }
                    }
                }
            }
        }
        
        /// <summary>
        /// Gets the agreement details to perform mail merge.
        /// </summary>
        public static List<AgreementDetails> GetAgreementDeatils()
        {
            List<AgreementDetails> details = new List<AgreementDetails>();
            details.Add(new AgreementDetails("LA-2026-10001", "March 15, 2026", "California", "Jennifer Anderson", "2458 Sunset Boulevard, Apt 3B", "Los Angeles", "California", "90028", "+1-213-555-0142", "jennifer.anderson@email.com", "Home Loan - Fixed Rate", "$450,000", "6.75", "360","$2,918", "$2,250", "Residential Property", "3-bedroom house at 2458 Sunset Blvd, Los Angeles, CA", "$585,000", "April 15, 2026", "April 15, 2056", "2% of outstanding principal", "$75", DateTime.Today.ToString(), DateTime.Today.ToString(), DateTime.Today.ToString()));
            details.Add(new AgreementDetails("LA-2026-10002", "March 10, 2026", "England", "James Mitchell", "47 Baker Street, Westminster", "London", "Greater London", "W1U 6TE", "+44-20-7946-0958", "james.mitchell@business.co.uk", "Business Expansion Loan", "£250,000", "7.50", "120", "£2,988", "£3,750", "Commercial Property", "Office space at 47 Baker Street, London", "£180,000", "April 10, 2026", "April 10, 2036", "3% of outstanding principal", "£50", DateTime.Today.ToString(), DateTime.Today.ToString(), DateTime.Today.ToString()));
            details.Add(new AgreementDetails("LA-2026-10003", "March 12, 2026", "New South Wales", "Sarah Williams", "156 George Street, Level 8", "Sydney", "New South Wales", "2000", "+61-2-9876-5432", "sarah.williams@email.com.au", "Vehicle Loan - New Car", "AU$65,000", "8.25", "60", "AU$1,324", "AU$975", "Motor Vehicle", "2026 Toyota Camry Hybrid", "AU$65,000", "April 12, 2026", "April 12, 2031", "No prepayment penalty", "AU$40", DateTime.Today.ToString(), DateTime.Today.ToString(), DateTime.Today.ToString()));
            details.Add(new AgreementDetails("LA-2026-10004", "March 8, 2026", "Ontario", "Michael Chen", "890 Bay Street, Suite 1205", "Toronto", "Ontario", "M5S 3A1", "+1-416-555-7890", "michael.chen@email.ca", "Personal Loan - Debt Consolidation", "CA$75,000", "9.50", "84", "CA$1,128", "CA$1,125", "Fixed Deposit", "Term deposit at TD Bank", "CA$30,000", "April 8, 2026", "April 8, 2033", "2.5% of outstanding principal", "CA$50", DateTime.Today.ToString(), DateTime.Today.ToString(), DateTime.Today.ToString()));
            details.Add(new AgreementDetails("LA-2026-10005", "March 18, 2026", "Bavaria", "Emma Schmidt", "Maximilianstraße 45", "Munich", "Bavaria", "80539", "+49-89-1234-5678", "emma.schmidt@email.de", "Investment Property Loan", "€380,000", "5.85", "300", "€2,247", "€5,700", "Investment Property", "2-bedroom apartment at Maximilianstraße 45, Munich", "€480,000", "April 18, 2026", "April 18, 2051", "1.5% of outstanding principal", "€60", DateTime.Today.ToString(), DateTime.Today.ToString(), DateTime.Today.ToString()));
            details.Add(new AgreementDetails("LA-2026-10006", "March 20, 2026", "Texas", "Robert Johnson", "1200 Commerce Street, Suite 400", "Dallas", "Texas", "75202", "+1-214-555-3456", "robert.johnson@startup.com", "Small Business Startup Loan", "$120,000", "10.25", "72", "$1,986", "$1,800", "Business Assets", "Office furniture, computers, and inventory", "$45,000", "April 20, 2026", "April 20, 2032", "3.5% of outstanding principal", "$85", DateTime.Today.ToString(), DateTime.Today.ToString(), DateTime.Today.ToString()));
            return details;
        }      
    }

    /// <summary>
    /// Represents a class to maintain agreement details.
    /// </summary>
    public class AgreementDetails
    {
        public string AgreementNumber { get; set; }
        public string AgreementDate { get; set; }
        public string State { get; set; }
        public string BorrowerName { get; set; }
        public string BorrowerAddress { get; set; }
        public string BorrowerCity { get; set; }
        public string BorrowerState { get; set; }
        public string BorrowerZip { get; set; }
        public string BorrowerPhone { get; set; }
        public string BorrowerEmail { get; set; }
        public string LoanProduct { get; set; }
        public string LoanAmount { get; set; }
        public string InterestRate { get; set; }
        public string LoanTerm { get; set; }
        public string MonthlyPayment { get; set; }
        public string ProcessingFee { get; set; }
        public string CollateralType { get; set; }
        public string CollateralDescription { get; set; }
        public string CollateralValue { get; set; }
        public string RepaymentStartDate { get; set; }
        public string MaturityDate { get; set; }
        public string PrepaymentPenalty { get; set; }
        public string LatePaymentFee { get; set; }
        public string WitnessName { get; set; }
        public string WitnessDate { get; set; }
        public string BorrowerDate { get; set; }
        public string LenderDate { get; set; }

        public AgreementDetails(string agreementNumber, string agreementDate, string state, string borrowerName, string borrowerAddress, string borrowerCity, string borrowerState, string borrowerZip, string borrowerPhone, string borrowerEmail, string loanProduct, string loanAmount,
            string interestRate, string loanTerm, string monthlyPayment, string processingFee, string collateralType, string collateralDescription, string collateralValue, string repaymentStartDate, string maturityDate, string prepaymentPenalty, string latePaymentFee, string witnessDate, string borrowerDate, string lenderDate)
        {
            AgreementNumber = agreementNumber;
            AgreementDate = agreementDate;
            State = state;
            BorrowerName = borrowerName;
            BorrowerAddress = borrowerAddress;
            BorrowerCity = borrowerCity;
            BorrowerState = borrowerState;
            BorrowerZip = borrowerZip;
            BorrowerPhone = borrowerPhone;
            BorrowerEmail = borrowerEmail;
            LoanProduct = loanProduct;
            LoanAmount = loanAmount;
            InterestRate = interestRate;
            LoanTerm = loanTerm;
            MonthlyPayment = monthlyPayment;
            ProcessingFee = processingFee;
            CollateralType = collateralType;
            CollateralDescription = collateralDescription;
            CollateralValue = collateralValue;
            RepaymentStartDate = repaymentStartDate;
            MaturityDate = maturityDate;
            PrepaymentPenalty = prepaymentPenalty;
            LatePaymentFee = latePaymentFee;
            WitnessDate = witnessDate;
            BorrowerDate = borrowerDate;
            LenderDate = lenderDate;
        }
    }
}


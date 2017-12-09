using System;
namespace IMD.Connect.SPO.Provisioning
{
    class IMDConnect
    {
        #region Properties
        public static string SiteUrl { get; set; }
        public static string ClientID { get; set; }
        public static string ClientSecrete { get; set; }
        public static string Action { get; set; }
        public static string APIKey { get; set; }
        public static string WorkflowID { get; set; }
        public static string ApiRootUrl { get; set; }
        public static string ExportPath { get; set; }
        public static string FilePath { get; set; }
        //public static string apiRootUrl { get; set; }
        //public static string FilePath { get; set; }
        #endregion

        static void Main(string[] args)
        {
            SiteUrl = "https://imdtst.sharepoint.com";
            ClientID = "bb07a71b-36bb-4070-812b-cd9fa7909459";
            ClientSecrete = "QFz7gj3lrx1hdvbo3Bj50wij7m5qaJ7Z0uE52HlJQEo=";
            FilePath = "D:\\DMS\\Projects\\PnPXMLTemplate.xml";

            //SiteUrl = "https://imdtst.sharepoint.com/crk";
            //APIKey = "5237477280d042db9122296c697bdb2c";
            //WorkflowID = "cff451cd-ea74-4255-bf70-ddcab358b780";
            //ExportPath = "D:/Nintex";
           // FilePath = "D:/Nintex/createdocLibrarychk.nwp";
           

            #region takinginput parameters
            //Console.WriteLine("Please provide Site Url:");
            //SiteUrl = Console.ReadLine();
            //Console.WriteLine("Please provide Client ID:");
            //ClientID = Console.ReadLine();
            //Console.WriteLine("Please provide Client Secrete:");
            //ClientSecrete = Console.ReadLine();
            #endregion

            Console.WriteLine("***************************");
            Console.WriteLine("Select 1 for Site Columns creation");
            Console.WriteLine("Select 2 for ContentTypes Creation and adding Site Columns");
            Console.WriteLine("Select 3 for Nintex Workflow Export");
            Console.WriteLine("Select 4 for Nintex Workflow Import");
            Console.WriteLine("Select 5 for Nintex Workflow Publish");
            Console.WriteLine("Select 6 for Nintex Form Export");
            Console.WriteLine("Select 7 for Nintex Form Import");
            Console.WriteLine("Select 8 for Nintex Form Publish");
            Console.WriteLine("***************************");
        
            Action = Console.ReadLine();
            switch (Action)
            {
                case "1":
                    //Console.WriteLine("Please provide xml file path");
                    //FilePath = Console.ReadLine();
                    Console.WriteLine("Creating Site Columns");
                    CreatingSiteColumns.SiteCoumnsCreation(FilePath);
                    Console.WriteLine("Site Columns Creation is Completed");
                    Console.ReadLine();
                    break;
                case "2":
                    Console.WriteLine("Creating Content Types and Adding Site Columns");
                    CreatingContentTypes.ContentTypeCreation(FilePath);
                    Console.WriteLine("Content type is completed");
                    Console.ReadLine();
                    break;
                case "3":
                    Console.WriteLine("Exporting Nintex Workflows");
                    ExportWorkflow.ExportWorkflowToFile();          
                    ExportWorkflow.UplodeWorkflow();
                    Console.WriteLine("Export Workflow is completed");
                    Console.ReadLine();

                    //Console.WriteLine("Please provide Nintex API Key for Export");
                    //APIKey = Console.ReadLine();
                    //Console.WriteLine("Please provide Workflow ID for Export");
                    //WorkflowID = Console.ReadLine();
                    //Console.WriteLine("Please provide Export Path");
                    //ExportPath = Console.ReadLine();
                    //Console.WriteLine("Please provide File Path");
                    //FilePath = Console.ReadLine();
                    ////ExportWorkflow.UplodeWorkflow();

                    break;
                case "4":
                    Console.WriteLine("Importing Nintex Workflow");
                    ImportWorkflow.CopyWorkflowToList();
                    Console.WriteLine("Import Workflow is completed");
                    Console.ReadKey();
                    break;
                case "5":
                    Console.WriteLine("Publishing Nintex Workflow");
                    PublishWorkflow.PublishingWorkflow();
                    Console.ReadKey();
                    break;
                case "6":
                    Console.WriteLine("Exporting Nintex Form");
                    Console.ReadKey();
                    break;
                case "7":
                    Console.WriteLine("Importing Nintex Form");
                    Console.ReadKey();
                    break;
                case "8":
                    Console.WriteLine("Publishing Nintex Form");
                    Console.ReadKey();
                    break;
                default: Console.WriteLine("You have selected Invalid Action");
                    Console.ReadKey();
                    break;
            }

        }
    }
}

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
        #endregion
        
        static void Main(string[] args)
        {
            //#region takinginput parameters
            Console.WriteLine("Please Provide OAuth Details");
            Console.WriteLine("Site Url:");
            SiteUrl = Console.ReadLine();
            Console.WriteLine("Client ID:");
            ClientID = Console.ReadLine();
            Console.WriteLine("Client Secret:");
            ClientSecrete = Console.ReadLine();
            //#endregion

            //SiteUrl = "https://imdtst.sharepoint.com/sites/KMS";
            //ClientID = "0437d24f-d85d-4817-b22e-ac02d5158fd0";
            //ClientSecrete = "QGKDpyhDAOknQfswg+Qi9O2TLq6h6rlCb+6qLtnA1qU=";
            //SiteUrl = "https://pavanisurya.sharepoint.com/sites/KMS";
            //ClientID = "d3c74aa7-a2fd-4a02-be15-93a75aa1f657";
            //ClientSecrete = "ub4VgkK4/Ou9W5jJlUYXmF0PiBFJqH5BBTTeV1bBCz4=";

            Console.WriteLine("**********************************************");
            Console.WriteLine("Type 1 for Term Store provisioning           *");
            Console.WriteLine("Type 2 for Lookup Lists Provisioning         *");
            Console.WriteLine("Type 3 for Site Columns provisioning         *");
            Console.WriteLine("Type 3 for Site Columns provisioning         *");
            Console.WriteLine("Type 4 for ContentTypes provisioning         *");
            Console.WriteLine("Type 5 for Master Lists provisioning         *");
            Console.WriteLine("Type 6 for Nintex Workflow Export            *");
            Console.WriteLine("Type 7 for Nintex Workflow Import            *");
            Console.WriteLine("Type 8 for Nintex Workflow Publish           *");
            Console.WriteLine("Type 9 for Nintex Form Export                *");
            Console.WriteLine("Type 10 for Nintex Form Import               *");
            Console.WriteLine("Type 11 for Nintex Form Publish              *");
            Console.WriteLine("Type 12 for Save as template                 *");
            Console.WriteLine("**********************************************");
            Console.Write("Please select option:");
            Action = Console.ReadLine();
            try
            {
                switch (Action)
                {
                    case "1":
                        //Creating Taxonomy 
                        Taxonomy.CreateTaxnomy();
                        break;
                    case "2":
                        //Creating Lookup Lists
                        Lists.CreateLists();
                        break;
                    case "3":
                        //Creating Site Columns
                        SiteColumns.CreateSiteColumns();
                        break;
                    case "4":
                        //Creating Content Types
                        ContentTypes.CreateContentTypes();
                        break;
                    case "5":
                        //Creating Master Lists
                        MasterLists.CreateMasterLists();
                        Console.WriteLine("Provisioning Lists");
                        break;
                    case "6":
                        ExportWorkflow.ExportWorkflowToFile();
                        ExportWorkflow.UploadWorkflow();
                        break;
                    case "7":
                        Console.WriteLine("Importing Nintex Workflow");
                        ImportWorkflow.CopyWorkflowToList();
                        Console.WriteLine("Import Workflow is completed");
                        Console.ReadKey();
                        break;
                    case "8":
                        Console.WriteLine("Publishing Nintex Workflow");
                        PublishWorkflow.PublishingWorkflow();
                        Console.ReadKey();
                        break;
                    case "9":
                        Console.WriteLine("Exporting Nintex Form");
                        ExportForm.ExportFormToFile();
                        Console.ReadKey();
                        break;
                    case "10":
                        Console.WriteLine("Importing Nintex Form");
                        ImportForm.CopyFormToList();
                        Console.ReadKey();
                        break;
                    case "11":
                        Console.WriteLine("Publishing Nintex Form");
                        PublishForm.PublishingForm();
                        Console.ReadKey();
                        break;
                    case "12":
                        Console.WriteLine("save as template");
                       // Lists.CreateList();
                        
                        Console.ReadKey();
                        break;
                    default:

                        Console.WriteLine("You have selected Invalid Action");
                        Console.ReadKey();
                        break;
                }

            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message.ToString());
                Console.ReadLine();
            }
            

        }
    }
}

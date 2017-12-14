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
            #region takinginput parameters
            //Console.WriteLine("Please provide Site Url:");
            //SiteUrl = Console.ReadLine();
            //Console.WriteLine("Please provide Client ID:");
            //ClientID = Console.ReadLine();
            //Console.WriteLine("Please provide Client Secrete:");
            //ClientSecrete = Console.ReadLine();
            #endregion
            Console.WriteLine("**********************************************************");
            Console.WriteLine("Option 1 for Site Columns creation");
            Console.WriteLine("Option 2 for ContentTypes Creation and adding Site Columns");
            Console.WriteLine("Option 3 for Nintex Workflow Export");
            Console.WriteLine("Option 4 for Nintex Workflow Import");
            Console.WriteLine("Option 5 for Nintex Workflow Publish");
            Console.WriteLine("Option 6 for Nintex Form Export");
            Console.WriteLine("Option 7 for Nintex Form Import");
            Console.WriteLine("Option 8 for Nintex Form Publish");
            Console.WriteLine("**********************************************************");
            Console.Write("Please select option:");
            Action = Console.ReadLine();
            try
            {
                switch (Action)
                {
                    case "1":
                      
                        CreatingSiteColumns.SiteCoumnsCreation();
                        Console.Write("*****Site Columns Creation is Completed.********");
                        Console.ReadLine();
                        break;
                    case "2":
                        
                        CreatingContentTypes.ContentTypeCreation();
                        Console.WriteLine("Content type is completed.");
                        Console.ReadLine();
                        break;
                    case "3":
                        Console.WriteLine("Exporting Nintex Workflows");
                        ExportWorkflow.ExportWorkflowToFile();
                        ExportWorkflow.UploadWorkflow();
                        Console.WriteLine("Export Workflow is completed");
                        Console.ReadLine();

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
                        ExportForm.ExportFormToFile();
                        Console.ReadKey();
                        break;
                    case "7":
                        Console.WriteLine("Importing Nintex Form");
                        ImportForm.CopyFormToList();
                        Console.ReadKey();
                        break;
                    case "8":
                        Console.WriteLine("Publishing Nintex Form");
                        PublishForm.PublishingForm();
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

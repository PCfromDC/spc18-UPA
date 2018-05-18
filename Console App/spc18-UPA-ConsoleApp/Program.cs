using System;
using System.Configuration;
using System.Data.SqlClient;
using System.IO;
using System.Security;
using System.Text;
using Microsoft.Online.SharePoint.TenantManagement;
using Microsoft.SharePoint.Client;

namespace spc18_UPA_ConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            #region Variables:
                string adminURL = "https://pcfromdc-admin.sharepoint.com";  // SPO Admin Site URL
                string siteURL = "https://pcfromdc.sharepoint.com/sites/spc18";  // Site where we upload JSON file
                string importFileURL = "https://pcfromdc.sharepoint.com/sites/spc18/upaSync/upaOutput-WebJob.txt"; 
                string docLibName = "UPA Sync"; // Document Library Name for upload
                string spoUserName = ConfigurationManager.AppSettings["spoUserName"];
                string spoPassword = ConfigurationManager.AppSettings["spoPassword"];
            #endregion

            #region Query SQL and Build JSON'ish String for upload to O365
                SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();
                builder.DataSource = "pcdemo.database.windows.net";
                builder.UserID = ConfigurationManager.AppSettings["dataBaseUserName"];
                builder.Password = ConfigurationManager.AppSettings["dataBasePW"];
                builder.InitialCatalog = "pcDemo_Personnel";
            #endregion

            #region Start to build jsonOutput string
                StringBuilder jsonSB = new StringBuilder();
                jsonSB.AppendLine("{");
                jsonSB.AppendLine("\"value\":");
                jsonSB.AppendLine("[");
            #endregion

            #region Get info from Azure SQL Table
                using (SqlConnection connection = new SqlConnection(builder.ConnectionString))
                {
                    connection.Open();
                    StringBuilder sb = new StringBuilder();
                    sb.Append("SELECT TOP(10) mail, city  ");
                    sb.Append("FROM pcDemo_SystemUsers ");
                    sb.Append("Where city is not null ");
                    sb.Append("and UserEmail like '%@pcfromdc.com' ");
                    sb.Append("or UserEmail like '%pcdemo.net'");
                    String sql = sb.ToString();

                    using (SqlCommand command = new SqlCommand(sql, connection))
                    {
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                jsonSB.AppendLine("{");
                                jsonSB.AppendLine("\"IdName\": \"" + reader.GetString(0) + "\",");
                                jsonSB.AppendLine("\"Property1\": \"" + reader.GetString(1) + "\"");
                                jsonSB.AppendLine("},");
                            }
                        }
                    }
                }
                Console.WriteLine("SQL query completed...");
            #endregion

            #region finish json'ish string and convert to stream
                // Clean up jsonSB and remove last comma
                string jasonClean = jsonSB.ToString();
                jasonClean = (jasonClean.Trim()).TrimEnd(',');
                // Add jasonClean back into StringBuilder
                StringBuilder jsonEnd = new StringBuilder(jasonClean);
                jsonEnd.AppendLine("");
                jsonEnd.AppendLine("]");
                jsonEnd.AppendLine("}");
                string jsonOutput = jsonEnd.ToString();
                Console.WriteLine("JSON build completed...");

                // Convert String to Stream
                byte[] byteArray = Encoding.UTF8.GetBytes(jsonOutput);
                MemoryStream stream = new MemoryStream(byteArray);
                Console.WriteLine("JSON converted...");
            #endregion

            #region Upload JSON file to SPO
                using (var clientContext = new ClientContext(siteURL))
                {
                    // set username and password
                    var passWord = new SecureString();
                    foreach (char c in spoPassword.ToCharArray()) passWord.AppendChar(c);
                    clientContext.Credentials = new SharePointOnlineCredentials(spoUserName, passWord);

                    Web web = clientContext.Web;
                    FileCreationInformation newFile = new FileCreationInformation();
                    newFile.Overwrite = true;
                    newFile.ContentStream = stream;
                    newFile.Url = importFileURL;
                    List docLibrary = web.Lists.GetByTitle(docLibName);
                    docLibrary.RootFolder.Files.Add(newFile);
                    clientContext.Load(docLibrary);
                    clientContext.ExecuteQuery();      
                }
                Console.WriteLine("File Uploaded...");
            #endregion

            #region Bulk Upload API
            using (var clientContext = new ClientContext(adminURL))
            {
                // set username and password
                var passWord = new SecureString();
                foreach (char c in spoPassword.ToCharArray()) passWord.AppendChar(c);
                clientContext.Credentials = new SharePointOnlineCredentials(spoUserName, passWord);

                // Get Tenant Context
                Office365Tenant tenant = new Office365Tenant(clientContext);
                clientContext.Load(tenant);
                clientContext.ExecuteQuery();

                // Only to check connection and permission, could be removed
                clientContext.Load(clientContext.Web);
                clientContext.ExecuteQuery();
                string title = clientContext.Web.Title;
                Console.WriteLine("Logged into " + title + "...");

                clientContext.Load(clientContext.Web);
                ImportProfilePropertiesUserIdType userIdType = ImportProfilePropertiesUserIdType.Email;
                var userLookupKey = "IdName";
                var propertyMap = new System.Collections.Generic.Dictionary<string, string>();
                propertyMap.Add("Property1", "City");
                // propertyMap.Add("Property2", "Office");
                var workItemId = tenant.QueueImportProfileProperties(userIdType, userLookupKey, propertyMap, importFileURL);
                clientContext.ExecuteQuery();
            }
            Console.WriteLine("UPA Bulk Update Completed...");
            #endregion
        }
    }
}
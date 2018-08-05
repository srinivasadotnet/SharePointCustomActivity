using Microsoft.SharePoint.Client;
using SarePointCustomActivities.Common;
using System.Activities;
using System.ComponentModel;
using System.Security;

namespace SarePointCustomActivities
{
    public class UploadDocumentToLibrary : CodeActivity
    {
        /// <summary>
        /// The SharePointSiteUri
        /// </summary>
        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> SharePointSiteUri { get; set; }

        /// <summary>
        /// The UserName
        /// </summary>
        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> UserName { get; set; }

        /// <summary>
        /// The Password
        /// </summary>
        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> Password { get; set; }

        /// <summary>
        /// The List Name
        /// </summary>
        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> ListName { get; set; }

        /// <summary>
        /// The Document Path
        /// </summary>
        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> FilePath { get; set; }

        /// <summary>
        /// The Document Path
        /// </summary>
        [Category("Input")]
        [RequiredArgument]
        public InArgument<bool> IsMultiFileUpload { get; set; }

        /// <summary>
        /// The ListItems
        /// </summary>
        [Category("Output")]
        public OutArgument<string> ResponseMessage { get; set; }

        /// <summary>
        /// The Execute
        /// </summary>
        /// <param name="context">Code Activity Context</param>
        protected override void Execute(CodeActivityContext context)
        {
            var userName = UserName.Get(context);
            var listName = ListName.Get(context);
            var filePath = FilePath.Get(context);
            var isMultiFileUpload = IsMultiFileUpload.Get(context);

            if (isMultiFileUpload ? System.IO.Directory.Exists(filePath) :System.IO.File.Exists(filePath))
            {
                var securePassword = new SecureString();

                foreach (char c in Password.Get(context))
                {
                    securePassword.AppendChar(c);
                }

                using (var clientContext = new ClientContext(SharePointSiteUri.Get(context)))
                {
                    clientContext.Credentials = new SharePointOnlineCredentials(userName, securePassword);
                    var listItemOperations = new ListItemOperations(clientContext, listName);

                    // Validate whether given List Context Type is Document Library or not
                    // If not Don't upload documents
                    if (listItemOperations.IsDocumentLibrary)
                    {
                        if (isMultiFileUpload)
                        {
                            ProcessDirectory(listItemOperations, filePath);
                        }
                        else
                        {
                            UploadFile(listItemOperations, filePath);
                        }
                        ResponseMessage.Set(context, $"Document Uploaded to {listName} Document Lirary");
                    }
                    else
                    {
                        ResponseMessage.Set(context, $"Given List {listName} is Not a Document Lirary");
                    }
                }
            }
            else
            {
                ResponseMessage.Set(context, $"File / File Path Not found");
            }
        }

        /// <summary>
        /// The ProcessDirectory
        /// </summary>
        /// <param name="listItemOperations"></param>
        /// <param name="targetDirectory"></param>
        public static void ProcessDirectory(ListItemOperations listItemOperations, string targetDirectory)
        {
            // Process the list of files found in the directory.
            string[] fileEntries = System.IO.Directory.GetFiles(targetDirectory);
            foreach (string fileName in fileEntries)
                UploadFile(listItemOperations, fileName);

            // Recurse into subdirectories of this directory.
            string[] subdirectoryEntries = System.IO.Directory.GetDirectories(targetDirectory);
            foreach (string subdirectory in subdirectoryEntries)
                ProcessDirectory(listItemOperations, subdirectory);
        }

        /// <summary>
        /// The UploadFile
        /// </summary>
        /// <param name="listItemOperations"></param>
        /// <param name="filePath"></param>
        public static void UploadFile(ListItemOperations listItemOperations, string filePath)
        {
            listItemOperations.UploadDocumentItem(filePath);
        }
    }
}

using Microsoft.SharePoint.Client;
using SarePointCustomActivities.Common;
using System;
using System.Activities;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Security;

namespace SarePointCustomActivities
{
    public class InsertListItems : CodeActivity
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
        /// The Input File Path
        /// </summary>
        [Category("Input")]
        [RequiredArgument]
        public InArgument<DataTable> InputData { get; set; }

        /// <summary>
        /// The Input File Path
        /// </summary>
        [Category("Output")]
        public OutArgument<string> ResultMessage { get; set; }

        /// <summary>
        /// The Execute
        /// </summary>
        /// <param name="context">Code Activity Context</param>
        protected override void Execute(CodeActivityContext context)
        {
            var userName = UserName.Get(context);
            var listName = ListName.Get(context);
            var inputData = InputData.Get(context);

            if (inputData.Rows.Count > 0)
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
                    var columnList = inputData.Columns;
                    foreach (DataRow dataRow in inputData.Rows)
                    {
                        Dictionary<string, string> newItemRecord = new Dictionary<string, string>();
                        foreach (var column in columnList)
                        {
                            newItemRecord.Add(column.ToString(), dataRow[column.ToString()].ToString());
                        }
                        if (newItemRecord.Count > 0)
                        {
                            listItemOperations.CreateListItem(newItemRecord);
                        }
                    }
                    ResultMessage.Set(context, $"Records Inserted to SharePoint List, Inserted Items : {inputData.Rows.Count}");
                }
            }
            else
            {
                ResultMessage.Set(context, "DataTable table is empty. Nothing to insert.");
            }
        }
    }
}

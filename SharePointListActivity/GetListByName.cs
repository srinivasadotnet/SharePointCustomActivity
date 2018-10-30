using Microsoft.SharePoint.Client;
using SharePointCustomActivities.Common;
using System;
using System.Activities;
using System.ComponentModel;
using System.Data;
using System.Security;

namespace SharePointCustomActivities
{
    public class GetListByName : CodeActivity
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
        /// The ListItems
        /// </summary>
        [Category("Output")]
        [RequiredArgument]
        public OutArgument<DataTable> ListItems { get; set; }

        /// <summary>
        /// The Execute
        /// </summary>
        /// <param name="context">Code Activity Context</param>
        protected override void Execute(CodeActivityContext context)
        {
            var userName = UserName.Get(context);
            var listName = ListName.Get(context);
            var securePassword = new SecureString();

            foreach (char c in Password.Get(context))
            {
                securePassword.AppendChar(c);
            }

            using (var clientContext = new ClientContext(SharePointSiteUri.Get(context)))
            {
                clientContext.Credentials = new SharePointOnlineCredentials(userName, securePassword);
                var listItemOperations = new ListItemOperations(clientContext, listName);
                ListItems.Set(context, DataHelperUtility.GetDataTableFromListItemCollection(listItemOperations.GetListItems()));
            }
        }
    }
}

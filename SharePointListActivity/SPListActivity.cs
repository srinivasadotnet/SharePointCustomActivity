using Microsoft.SharePoint.Client;
using System.Activities;
using System.ComponentModel;
using System.Security;
using SPListCustomActivity.Common;

namespace SPListCustomActivity
{
    public class SPListActivity : CodeActivity
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
        /// The SPAvailableLists
        /// </summary>
        [Category("Output")]
        public OutArgument<ListCollection> SPAvailableLists { get; set; }

        /// <summary>
        /// The Execute
        /// </summary>
        /// <param name="context">Code Activity Context</param>
        protected override void Execute(CodeActivityContext context)
        {
            var userName = UserName.Get(context);
            var securePassword = new SecureString();

            foreach (char c in Password.Get(context))
            {
                securePassword.AppendChar(c);
            }

            using (var clientContext = new ClientContext(SharePointSiteUri.Get(context)))
            {
                clientContext.Credentials = new SharePointOnlineCredentials(userName, securePassword);
                var listItemOperations = new ListItemOperations(clientContext, null);
                SPAvailableLists.Set(context, listItemOperations.GetListCollection());
            }
        }
    }
}

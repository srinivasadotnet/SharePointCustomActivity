using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace SharePointCustomActivities.Common
{
    public class ListItemOperations : IListOperations
    {
        /// <summary>
        /// The clientContext
        /// </summary>
        ClientContext clientContext;

        /// <summary>
        /// The list
        /// </summary>
        List list;

        /// <summary>
        /// The IsDocumentLibrary
        /// </summary>S
        public bool IsDocumentLibrary
        {
            get
            {
                clientContext.Load(list, li => li.BaseType);
                clientContext.ExecuteQuery();
                return list.BaseType == BaseType.DocumentLibrary;
            }
        }

        /// <summary>
        /// The ListItemOperations Constructor 
        /// </summary>
        /// <param name="clientContext">The clientContext</param>
        /// <param name="taskListName">The listName</param>
        public ListItemOperations(ClientContext clientContext, string listName)
        {
            this.clientContext = clientContext;
            if (!string.IsNullOrWhiteSpace(listName))
                this.list = clientContext.Web.Lists.GetByTitle(listName);
        }

        /// <summary>
        /// The CreateListItem
        /// </summary>
        /// <param name="listItem">The listItem</param>
        public void CreateListItem(Dictionary<string, string> listItemValues)
        {
            ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
            ListItem listItem = list.AddItem(itemCreateInfo);
            listItem = SetListItemModel(listItem, listItemValues);
            listItem.Update();
            clientContext.ExecuteQuery();
        }

        /// <summary>
        /// The CreateListItem
        /// </summary>
        /// <param name="listItem">The listItem</param>
        public void UploadDocumentItem(string sourceFilePath)
        {
            clientContext.Load(list.RootFolder);
            clientContext.ExecuteQuery();
            var targetFileUrl = String.Format("{0}/{1}", list.RootFolder.ServerRelativeUrl.ToString(), Path.GetFileName(sourceFilePath));
            using (var fileStream = new FileStream(sourceFilePath, FileMode.Open))
            {
                Microsoft.SharePoint.Client.File.SaveBinaryDirect(clientContext, targetFileUrl, fileStream, true);
            }
            clientContext.ExecuteQuery();
        }

        /// <summary>
        /// The DeleteListItem
        /// </summary>
        /// <param name="itemId">The itemId</param>
        public void DeleteListItem(int itemId)
        {
            ListItem oListItem = list.GetItemById(itemId);
            oListItem.DeleteObject();
            clientContext.ExecuteQuery();
        }

        /// <summary>
        /// The GetListCollection
        /// </summary>
        /// <returns>List of Collections</returns>
        public ListCollection GetListCollection()
        {
            var web = clientContext.Web;
            clientContext.Load(web.Lists, lists => lists.Where(list => list.Hidden == false));
            clientContext.ExecuteQuery();
            return web.Lists;
        }

        /// <summary>
        /// The GetListItems
        /// </summary>
        /// <returns>The ListItemCollection</returns>
        public ListItemCollection GetListItems()
        {
            Microsoft.SharePoint.Client.CamlQuery camlQuery = new CamlQuery();
            ListItemCollection collListItem = list.GetItems(camlQuery);
            clientContext.Load(collListItem);
            clientContext.ExecuteQuery();
            return collListItem;
        }

        /// <summary>
        /// The GetListItemsById
        /// </summary>
        /// <param name="itemId">The itemId</param>
        /// <returns>The ListItem</returns>
        public ListItem GetListItemsById(int itemId)
        {
            return list.GetItemById(itemId);
        }

        /// <summary>
        /// The UpdateListItem
        /// </summary>
        /// <param name="listItemValues">The listItemValues</param>
        public void UpdateListItem(Dictionary<string, string> listItemValues)
        {
            ListItem listItem = list.GetItemById(listItemValues["Id"]);
            listItem = SetListItemModel(listItem, listItemValues);
            listItem.Update();
            clientContext.ExecuteQuery();
        }

        /// <summary>
        /// The SetListItemModel
        /// </summary>
        /// <param name="listItem">The listItem</param>
        /// <param name="listValues">The listValues</param>
        /// <returns></returns>
        private ListItem SetListItemModel(ListItem listItem, Dictionary<string, string> listValues)
        {
            foreach (var item in listValues)
            {
                listItem[item.Key] = item.Value;
            }

            return listItem;
        }
    }
}

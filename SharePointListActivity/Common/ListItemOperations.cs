using Microsoft.SharePoint.Client;
using System.Collections.Generic;
using System.Linq;

namespace SPListCustomActivity.Common
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
            CamlQuery camlQuery = new CamlQuery();
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

using Microsoft.SharePoint.Client;
using System.Collections.Generic;

namespace SPListCustomActivity.Common
{
    public interface IListOperations
    {
        /// <summary>
        /// The CreateListItem
        /// </summary>
        /// <param name="listItemValues"></param>
        void CreateListItem(Dictionary<string, string> listItemValues);

        /// <summary>
        /// The UpdateListItem
        /// </summary>
        /// <param name="listItemValues"></param>
        void UpdateListItem(Dictionary<string, string> listItemValues);

        /// <summary>
        /// The DeleteListItem
        /// </summary>
        /// <param name="itemId"></param>
        void DeleteListItem(int itemId);

        /// <summary>
        /// The GetListItems
        /// </summary>
        /// <returns></returns>
        ListItemCollection GetListItems();

        /// <summary>
        /// The GetListCollection
        /// </summary>
        /// <returns></returns>
        ListCollection GetListCollection();

        /// <summary>
        /// The GetListItemsById
        /// </summary>
        /// <param name="itemId"></param>
        /// <returns></returns>
        ListItem GetListItemsById(int itemId);
    }
}

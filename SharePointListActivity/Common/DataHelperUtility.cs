using Microsoft.SharePoint.Client;
using System;
using System.Data;

namespace SharePointCustomActivities.Common
{
    public static class DataHelperUtility
    {
        /// <summary>
        /// The GetDataTableFromListItemCollection
        /// </summary>
        /// <param name="listItems"></param>
        /// <returns></returns>
        public static DataTable GetDataTableFromListItemCollection(ListItemCollection listItems)
        {
            string strWhere = string.Empty;
            DataTable dtGetReqForm = new DataTable();

            try
            {
                if (listItems != null && listItems.Count > 0)
                {
                    foreach (var field in listItems[0].FieldValues.Keys)
                    {
                        dtGetReqForm.Columns.Add(field);
                    }

                    foreach (var item in listItems)
                    {
                        DataRow dr = dtGetReqForm.NewRow();

                        foreach (var obj in item.FieldValues)
                        {
                            if (obj.Value != null)
                            {
                                string key = obj.Key;
                                string type = obj.Value.GetType().FullName;

                                if (type == "Microsoft.SharePoint.Client.FieldLookupValue")
                                {
                                    dr[obj.Key] = ((FieldLookupValue)obj.Value).LookupValue;
                                }
                                else if (type == "Microsoft.SharePoint.Client.FieldUserValue")
                                {
                                    dr[obj.Key] = ((FieldUserValue)obj.Value).LookupValue;
                                }
                                else if (type == "Microsoft.SharePoint.Client.FieldUserValue[]")
                                {
                                    FieldUserValue[] multValue = (FieldUserValue[])obj.Value;
                                    foreach (FieldUserValue fieldUserValue in multValue)
                                    {
                                        dr[obj.Key] += (fieldUserValue).LookupValue;
                                    }
                                }
                                else if (type == "System.DateTime")
                                {
                                    if (obj.Value.ToString().Length > 0)
                                    {
                                        var date = obj.Value.ToString().Split(' ');
                                        if (date[0].Length > 0)
                                        {
                                            dr[obj.Key] = date[0];
                                        }
                                    }
                                }
                                else
                                {
                                    dr[obj.Key] = obj.Value;
                                }
                            }
                            else
                            {
                                dr[obj.Key] = null;
                            }
                        }
                        dtGetReqForm.Rows.Add(dr);
                    }
                }

            }
            catch (Exception ex)
            {
            }
            finally
            {
            }

            return dtGetReqForm;
        }
    }
}

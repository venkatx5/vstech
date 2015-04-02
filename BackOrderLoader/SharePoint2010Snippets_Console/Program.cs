using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint.Client;
using Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.Threading;

namespace BackOrderLoader
{
    //Allocations "AllocationsTemp" "C:\Venkat\CommIT\Alloc\Allocation_Temp.xlsx"
    class Program
    {
        static Int64 intItemCount;
        static Dictionary<string, string> dictNutriTM, dictPharmacyTM, dictPortTM;
        static string strSiteURL = "https://worksites.connect.inbaxter.com/sites/USMDMO/ProductAllocations/";

        static void Main(string[] args)
        {
            string strModule = "",
                strListName = "",
                strExcelFilePath = "";
            intItemCount = 0;

            if (args.Length > 1)
            {
                strModule = args[0];
                strListName = args[1];
                strExcelFilePath = args[2];

                if(args.Length > 3)
                {
                    strSiteURL = args[3];
                }

                if (strModule.ToLower() == "backorder")
                {
                    Console.WriteLine("Processing BackOrders. File : {0}", strExcelFilePath);
                    intItemCount = ProcessBackOrders(strListName, strExcelFilePath);
                }
                else
                {
                    if (strModule.ToLower() == "customer")
                    {
                        Console.WriteLine("Processing Customers File. File : {0}", strExcelFilePath);
                        intItemCount = ProcessCustomers(strListName, strExcelFilePath);
                    }
                    else
                    {
                        Console.WriteLine("Processing Allocations. File : {0}", strExcelFilePath);
                        intItemCount = ProcessAllocations(strListName, strExcelFilePath);
                    }
                }

                Console.WriteLine("[{0}] Items updated on the list.", intItemCount);
            }
            else
            {
                Console.WriteLine("This application requires Parameters to proceed.");
            }

            
            Console.ReadKey();
       }


        /// <summary>
        /// Process Back Orders Data Load
        /// </summary>
        /// <param name="strListName">Target SharePoint List Name</param>
        /// <param name="strExcelFilePath">Source Excel File Path</param>
        /// <returns>No of items updated</returns>
        private static Int64 ProcessBackOrders(string strListName, string strExcelFilePath)
        {
            Int64 intNoOfItemsUpdated = 0;
            string strTMListName = "Reps Postal Code Master";

            using (ClientContext oClientContext = new ClientContext(strSiteURL))
            {
                Web oWeb = oClientContext.Web;
                if (oWeb == null)
                {
                    Console.WriteLine("Can't open the Site " + strSiteURL + ".\n Please check with Site Administrator.");
                }
                else
                {
                    Console.WriteLine("Clearing List [{0}] Started.", strListName);
                    //Delete All Items from Back Order List.
                    DeleteAllItems(oClientContext, strListName);
                    Console.WriteLine("Clearing List [{0}] Completed.\n\n", strListName);

                    Console.WriteLine("Loading TM List Started.");
                    LoadTMLists(oClientContext, strTMListName);
                    Console.WriteLine("Loading TM List Completed.");

                    Console.WriteLine("Upload Data Started.");
                    intNoOfItemsUpdated = AddNewBackOrderItem(oClientContext, oWeb, strListName, strExcelFilePath);
                    Console.WriteLine("Upload Data Completed.");
                }
            }

            return intNoOfItemsUpdated;
        }

        /// <summary>
        /// Process Allocations Data Load
        /// </summary>
        /// <param name="strListName">Target SharePoint List Name</param>
        /// <param name="strExcelFilePath">Source Excel File Path</param>
        /// <returns>No of items updated</returns>
        private static Int64 ProcessAllocations(string strListName, string strExcelFilePath)
        {
            Int64 intNoOfItemsUpdated = 0;

            using (ClientContext oClientContext = new ClientContext(strSiteURL))
            {
                oClientContext.Credentials = System.Net.CredentialCache.DefaultNetworkCredentials;

                Web oWeb = oClientContext.Web;
                if (oWeb == null)
                {
                    Console.WriteLine("Can't open the Site " + strSiteURL + ".\n Please check with Site Administrator.");
                }
                else
                {
                    Console.WriteLine("Clearing List [{0}] Started.", strListName);
                    //Delete All Items from Back Order List.
                    DeleteAllItems(oClientContext, strListName);
                    Console.WriteLine("Clearing List [{0}] Completed.\n\n", strListName);

                    Console.WriteLine("Upload Data Started.");
                    intNoOfItemsUpdated = AddNewAllocationItem(oClientContext, oWeb, strListName, strExcelFilePath);
                    Console.WriteLine("Upload Data Completed.");
                }
            }

            return intNoOfItemsUpdated;
        }

        /// <summary>
        /// Process Back Orders Data Load
        /// </summary>
        /// <param name="strListName">Target SharePoint List Name</param>
        /// <param name="strExcelFilePath">Source Excel File Path</param>
        /// <returns>No of items updated</returns>
        private static Int64 ProcessCustomers(string strListName, string strExcelFilePath)
        {
            Int64 intNoOfItemsUpdated = 0;
            

            using (ClientContext oClientContext = new ClientContext(strSiteURL))
            {
                Web oWeb = oClientContext.Web;
                if (oWeb == null)
                {
                    Console.WriteLine("Can't open the Site " + strSiteURL + ".\n Please check with Site Administrator.");
                }
                else
                {
                    Console.WriteLine("Clearing List [{0}] Started.", strListName);
                    ////Delete All Items from Back Order List.
                    //DeleteAllItems(oClientContext, strListName);
                    DeleteOldCustomers(oClientContext, strListName, strExcelFilePath);
                    Console.WriteLine("Clearing List [{0}] Completed.\n\n", strListName);
                    //ListAllItems(oClientContext, strListName, strExcelFilePath);

                    //Console.WriteLine("Upload Data Started.");
                    //intNoOfItemsUpdated = AddNewBackOrderItem(oClientContext, oWeb, strListName, strExcelFilePath);
                    //Console.WriteLine("Upload Data Completed.");
                }
            }

            return intNoOfItemsUpdated;
        }

        /// <summary>
        /// Delete All Items from given List.
        /// </summary>
        /// <param name="oClientContext">Client Context</param>
        /// <param name="strListName">List Name</param>
        /// <returns>Delete Result</returns>
        private static string DeleteAllItems(ClientContext oClientContext, string strListName)
        {
            string strDeleteResult = "Init";
            ListItemCollectionPosition licp = null;            
            Web oWeb = oClientContext.Web;

            List oList = oWeb.Lists.GetByTitle(strListName);

            oClientContext.Load(oList);


            DeleteDataThread objUploadDataThread = new DeleteDataThread();

            Thread[] thUploadDataThread = new Thread[10];
                        
            var intTempThreadCounter = 0;
            const int NoOfThreadsLimit = 10;

            ManualResetEvent[] doneEvents = new ManualResetEvent[NoOfThreadsLimit];
            DeleteDataThread[] aryUDThread = new DeleteDataThread[NoOfThreadsLimit];

            while(true)
            {
                CamlQuery oQuery = new CamlQuery();
                oQuery.ViewXml = @"<View><ViewFields><FieldRef Name='Id'/></ViewFields><RowLimit>250</RowLimit></View>";
                oQuery.ListItemCollectionPosition = licp;
                ListItemCollection oItems = oList.GetItems(oQuery);

                oClientContext.Load(oItems);
                oClientContext.ExecuteQuery();

                licp = oItems.ListItemCollectionPosition;

                doneEvents[intTempThreadCounter] = new ManualResetEvent(false);
                DeleteDataThread objUDT = new DeleteDataThread(oClientContext, oItems, doneEvents[intTempThreadCounter]);
                aryUDThread[intTempThreadCounter] = objUDT;
                ThreadPool.QueueUserWorkItem(objUDT.ThreadPoolCallback, intTempThreadCounter);

                intTempThreadCounter++;

                if (intTempThreadCounter > 9)
                {
                    foreach (var e in doneEvents)
                    {
                        if (e != null)
                        {
                            e.WaitOne();
                        }
                    } 
                    //WaitHandle.WaitAll(doneEvents);
                    Console.WriteLine("Delete Thread Reset. ");
                    intTempThreadCounter = 0;
                }

                if(licp == null)
                {
                    if (intTempThreadCounter > 0)
                    {
                        Console.WriteLine("Delete Thread Count at end " + intTempThreadCounter);
                        foreach (var e in doneEvents)
                        {
                            if (e != null)
                            {
                                e.WaitOne();
                            }
                        }
                        //for (int i = intTempThreadCounter; i < NoOfThreadsLimit; i++)
                        //{
                        //    ListItemCollection oItemsTemp = null;
                        //    doneEvents[i] = new ManualResetEvent(false);
                        //    DeleteDataThread objUDTTemp = new DeleteDataThread(oClientContext, oItemsTemp, doneEvents[i]);
                        //    aryUDThread[i] = objUDT;
                        //    ThreadPool.QueueUserWorkItem(objUDT.ThreadPoolCallback, i);
                        //}
                        //    WaitHandle.WaitAll(doneEvents);
                    }
                    Console.WriteLine("List [{0}] purged successfully.", strListName);
                    break;
                }
            }

            strDeleteResult = "Success";

            return strDeleteResult;
        }

        // <summary>
        /// Delete All Items from given List.
        /// </summary>
        /// <param name="oClientContext">Client Context</param>
        /// <param name="strListName">List Name</param>
        /// <param name="strExcelFilePath">Excel File Path</param>
        /// <returns>Delete Result</returns>
        private static string ListAllItems(ClientContext oClientContext, string strListName, string strExcelFilePath)
        {
            string strDeleteResult = "Init";
            ListItemCollectionPosition licp = null;
            Web oWeb = oClientContext.Web;

            List oList = oWeb.Lists.GetByTitle(strListName);

            oClientContext.Load(oList);
            
            //DeleteDataThread objUploadDataThread = new DeleteDataThread();

            //Thread[] thUploadDataThread = new Thread[10];

            //var intTempThreadCounter = 0;
            const int NoOfThreadsLimit = 10;

            //ManualResetEvent[] doneEvents = new ManualResetEvent[NoOfThreadsLimit];
            //DeleteDataThread[] aryUDThread = new DeleteDataThread[NoOfThreadsLimit];

            // Uncomment to Find Mismatch
            Application ExcelObj = new Application();

            Workbook theWorkbook = ExcelObj.Workbooks.Open(strExcelFilePath, 0, true, 5,
                     "", "", true, XlPlatform.xlWindows, "\t", false, false,
                     0, true);
            Sheets sheets = theWorkbook.Worksheets;
            Worksheet worksheet = (Worksheet)sheets.get_Item(1);

            bool isMoreItemToAdd = true;
            Int32 intItemIndex = 2;
            List<string> lstCustomerNo = new List<string>();

            while (isMoreItemToAdd == true)
            {
                Range range = worksheet.get_Range("A" + intItemIndex.ToString()); //, "B" + intItemIndex.ToString()

                if (range.Cells.Value == null)
                {
                    isMoreItemToAdd = false;
                    //Console.WriteLine("No. of Items. {0}", intItemIndex - 2);
                }
                else
                {
                    string strCustomerNo = range.Cells.Value.ToString();
                    //Console.Write(strCustomerNo);
                    lstCustomerNo.Add(strCustomerNo);                    
                }
                intItemIndex++;
            }

            theWorkbook.Close();
            Console.WriteLine("No of Customer Entries : {0}", lstCustomerNo.Count);
            
            //CamlQuery oQuery = new CamlQuery();            
            //do
            //{
            //    oQuery.ViewXml = @"<View>"                            
            //                + "<ViewFields><FieldRef Name='Customer_x0020_No' /></ViewFields>"
            //                + "<RowLimit Paged='TRUE'>5</RowLimit>"
            //                + "</View>";
              //+ "<Query>"
              //              + "<Where><Eq><FieldRef Name='Customer_x0020_No' /><Value Type='Text'>34133712</Value></Eq></Where>"
              //              + "</Query>"

            // <RowLimit>700000</RowLimit>
            //<Where><BeginsWith><FieldRef Name='Customer_x0020_No' /><Value Type='Text'>4000</Value></BeginsWith></Where>
            //  // do something with the page result

            //  // set the position cursor for the next iteration
            //  query.ListItemCollectionPosition = items.ListItemCollectionPosition;
            //} while (query.ListItemCollectionPosition != null)

            Int32 intNoOfItemsProcessed = 0, intNoOfItemsDeleted = 0;
            CamlQuery oQuery = new CamlQuery();

            List<string> lstCustomerNoExtra = new List<string>();

            do
            {
                oQuery.ViewXml = @"<View>"
                            + "<ViewFields><FieldRef Name='Customer_x0020_No'/></ViewFields>"
                            + "<RowLimit Paged='TRUE'>200</RowLimit>"
                            + "</View>";

                oQuery.ListItemCollectionPosition = licp;

                ListItemCollection oItems = oList.GetItems(oQuery);

                oClientContext.Load(oItems);
                oClientContext.ExecuteQuery();

                licp = oItems.ListItemCollectionPosition;
                foreach (ListItem oListItem in oItems)
                {
                    string strCustNo = oListItem["Customer_x0020_No"].ToString();
                    //lstCustomerNoExtra.Add(strCustNo);
                    intNoOfItemsProcessed++;
                    if (IsItemExists(lstCustomerNo, strCustNo) == true)
                    {
                        //lstCustomerNoExtra.Add(strCustNo);
                        //oListItem.DeleteObject();
                        lstCustomerNo.Remove(strCustNo);
                       // Console.Write(strCustNo + "|");
                        intNoOfItemsDeleted++;
                    }
                }

                //for (int intTempCounter = oItems.Count - 1; intTempCounter >= 0; intTempCounter--)
                //{
                //    ListItem oListItem = oItems[intTempCounter];
                //    string strCustNo = oListItem["Customer_x0020_No"].ToString();

                //    if (IsItemExists(lstCustomerNo, strCustNo) == true)
                //    {
                //        //lstCustomerNo.Remove(strCustNo);
                //       // oListItem.DeleteObject();
                //        //Console.Write(strCustNo + "|");
                //        Console.WriteLine(strCustNo);
                //        intNoOfItemsDeleted++;
                //    }
                //    intNoOfItemsProcessed++;
                //}

            } while (licp != null);

            //Console.WriteLine("Extra Entries : ");

            //var duplicateKeys = lstCustomerNoExtra.GroupBy(x => x)
            //            .Where(group => group.Count() > 1)
            //            .Select(group => group.Key);

            //foreach (string strCustNoTemp in duplicateKeys)
            //{
            //    Console.Write(strCustNoTemp + " | ");
            //    intNoOfItemsDeleted++;
            //}

            Console.WriteLine("Missing Customer No. ");
            foreach (string strCustNoTemp in lstCustomerNo)
            {
                Console.Write(strCustNoTemp + " | ");
                intNoOfItemsDeleted++;
            }

            Console.WriteLine("No of Entries Processed : {0} | Deleted : {1}", intNoOfItemsProcessed, intNoOfItemsDeleted);
            strDeleteResult = "Success";

            return strDeleteResult;
        }

        // <summary>
        /// Delete Old Customers on SharePoint List (Source : Excel File).
        /// </summary>
        /// <param name="oClientContext">Client Context</param>
        /// <param name="strListName">List Name</param>
        /// <param name="strExcelFilePath">Excel File Path</param>
        /// <returns>Delete Result</returns>
        private static string DeleteOldCustomers(ClientContext oClientContext, string strListName, string strExcelFilePath)
        {
            string strDeleteResult = "Init";
            ListItemCollectionPosition licp = null;
            Web oWeb = oClientContext.Web;

            List oList = oWeb.Lists.GetByTitle(strListName);

            oClientContext.Load(oList);

            // Uncomment to Find Mismatch
            Application ExcelObj = new Application();

            Workbook theWorkbook = ExcelObj.Workbooks.Open(strExcelFilePath, 0, true, 5,
                     "", "", true, XlPlatform.xlWindows, "\t", false, false,
                     0, true);
            Sheets sheets = theWorkbook.Worksheets;
            Worksheet worksheet = (Worksheet)sheets.get_Item(1);

            bool isMoreItemToAdd = true;
            Int32 intItemIndex = 2;
            List<string> lstCustomerNo = new List<string>();

            while (isMoreItemToAdd == true)
            {
                Range range = worksheet.get_Range("A" + intItemIndex.ToString()); //, "B" + intItemIndex.ToString()

                if (range.Cells.Value == null)
                {
                    isMoreItemToAdd = false;                    
                }
                else
                {
                    string strCustomerNo = range.Cells.Value.ToString();
                    
                    lstCustomerNo.Add(strCustomerNo);
                }
                intItemIndex++;
            }

            theWorkbook.Close();
            Console.WriteLine("No of Customer Entries to Delete : {0}", lstCustomerNo.Count);

            Int32 intNoOfItemsProcessed = 0, intNoOfItemsDeleted = 0;
            CamlQuery oQuery = new CamlQuery();

            List<string> lstCustomerNoExtra = new List<string>();

            do
            {
                oQuery.ViewXml = @"<View>"
                            + "<ViewFields><FieldRef Name='Customer_x0020_No'/></ViewFields>"
                            + "<RowLimit Paged='TRUE'>200</RowLimit>"
                            + "</View>";

                oQuery.ListItemCollectionPosition = licp;

                ListItemCollection oItems = oList.GetItems(oQuery);

                oClientContext.Load(oItems);
                oClientContext.ExecuteQuery();

                licp = oItems.ListItemCollectionPosition;                

                for (int intTempCounter = oItems.Count - 1; intTempCounter >= 0; intTempCounter--)
                {
                    ListItem oListItem = oItems[intTempCounter];
                    string strCustNo = oListItem["Customer_x0020_No"].ToString();

                    if (IsItemExists(lstCustomerNo, strCustNo) == true)
                    {
                        //lstCustomerNo.Remove(strCustNo);
                         oListItem.DeleteObject();
                         Console.Write("-");
                        //Console.Write(strCustNo + "|");                        
                        intNoOfItemsDeleted++;
                    }
                    intNoOfItemsProcessed++;
                }

            } while (licp != null);

            Console.WriteLine("No of Entries Processed : {0} | Deleted : {1}", intNoOfItemsProcessed, intNoOfItemsDeleted);
            strDeleteResult = "Success";

            return strDeleteResult;
        }

        // <summary>
        /// List all Customer Data missing in SharePoint List.
        /// </summary>
        /// <param name="oClientContext">Client Context</param>
        /// <param name="strListName">List Name</param>
        /// <param name="strExcelFilePath">Excel File Path</param>
        /// <returns>Delete Result</returns>
        private static string ListMissingCustomers(ClientContext oClientContext, string strListName, string strExcelFilePath)
        {
            string strDeleteResult = "Init";
            ListItemCollectionPosition licp = null;
            Web oWeb = oClientContext.Web;

            List oList = oWeb.Lists.GetByTitle(strListName);

            oClientContext.Load(oList);

            //var intTempThreadCounter = 0;
            const int NoOfThreadsLimit = 10;

            // Uncomment to Find Mismatch
            Application ExcelObj = new Application();

            Workbook theWorkbook = ExcelObj.Workbooks.Open(strExcelFilePath, 0, true, 5,
                     "", "", true, XlPlatform.xlWindows, "\t", false, false,
                     0, true);
            Sheets sheets = theWorkbook.Worksheets;
            Worksheet worksheet = (Worksheet)sheets.get_Item(1);

            bool isMoreItemToAdd = true;
            Int32 intItemIndex = 2;
            List<string> lstCustomerNo = new List<string>();

            while (isMoreItemToAdd == true)
            {
                Range range = worksheet.get_Range("A" + intItemIndex.ToString()); //, "B" + intItemIndex.ToString()

                if (range.Cells.Value == null)
                {
                    isMoreItemToAdd = false;
                    //Console.WriteLine("No. of Items. {0}", intItemIndex - 2);
                }
                else
                {
                    string strCustomerNo = range.Cells.Value.ToString();
                    //Console.Write(strCustomerNo);
                    lstCustomerNo.Add(strCustomerNo);
                }
                intItemIndex++;
            }

            theWorkbook.Close();
            Console.WriteLine("No of Customer Entries : {0}", lstCustomerNo.Count);

            //CamlQuery oQuery = new CamlQuery();            
            //do
            //{
            //    oQuery.ViewXml = @"<View>"                            
            //                + "<ViewFields><FieldRef Name='Customer_x0020_No' /></ViewFields>"
            //                + "<RowLimit Paged='TRUE'>5</RowLimit>"
            //                + "</View>";
            //+ "<Query>"
            //              + "<Where><Eq><FieldRef Name='Customer_x0020_No' /><Value Type='Text'>34133712</Value></Eq></Where>"
            //              + "</Query>"

            // <RowLimit>700000</RowLimit>
            //<Where><BeginsWith><FieldRef Name='Customer_x0020_No' /><Value Type='Text'>4000</Value></BeginsWith></Where>
            //  // do something with the page result

            //  // set the position cursor for the next iteration
            //  query.ListItemCollectionPosition = items.ListItemCollectionPosition;
            //} while (query.ListItemCollectionPosition != null)

            Int32 intNoOfItemsProcessed = 0, intNoOfItemsDeleted = 0;
            CamlQuery oQuery = new CamlQuery();

            List<string> lstCustomerNoExtra = new List<string>();

            do
            {
                oQuery.ViewXml = @"<View>"
                            + "<ViewFields><FieldRef Name='Customer_x0020_No'/></ViewFields>"
                            + "<RowLimit Paged='TRUE'>200</RowLimit>"
                            + "</View>";

                oQuery.ListItemCollectionPosition = licp;

                ListItemCollection oItems = oList.GetItems(oQuery);

                oClientContext.Load(oItems);
                oClientContext.ExecuteQuery();

                licp = oItems.ListItemCollectionPosition;
                foreach (ListItem oListItem in oItems)
                {
                    string strCustNo = oListItem["Customer_x0020_No"].ToString();
                    //lstCustomerNoExtra.Add(strCustNo);
                    intNoOfItemsProcessed++;
                    if (IsItemExists(lstCustomerNo, strCustNo) == true)
                    {
                        //lstCustomerNoExtra.Add(strCustNo);
                        //oListItem.DeleteObject();
                        lstCustomerNo.Remove(strCustNo);
                        // Console.Write(strCustNo + "|");
                        intNoOfItemsDeleted++;
                    }
                }

                //for (int intTempCounter = oItems.Count - 1; intTempCounter >= 0; intTempCounter--)
                //{
                //    ListItem oListItem = oItems[intTempCounter];
                //    string strCustNo = oListItem["Customer_x0020_No"].ToString();

                //    if (IsItemExists(lstCustomerNo, strCustNo) == true)
                //    {
                //        //lstCustomerNo.Remove(strCustNo);
                //       // oListItem.DeleteObject();
                //        //Console.Write(strCustNo + "|");
                //        Console.WriteLine(strCustNo);
                //        intNoOfItemsDeleted++;
                //    }
                //    intNoOfItemsProcessed++;
                //}

            } while (licp != null);

            //Console.WriteLine("Extra Entries : ");

            //var duplicateKeys = lstCustomerNoExtra.GroupBy(x => x)
            //            .Where(group => group.Count() > 1)
            //            .Select(group => group.Key);

            //foreach (string strCustNoTemp in duplicateKeys)
            //{
            //    Console.Write(strCustNoTemp + " | ");
            //    intNoOfItemsDeleted++;
            //}

            Console.WriteLine("Missing Customer No. ");
            foreach (string strCustNoTemp in lstCustomerNo)
            {
                Console.Write(strCustNoTemp + " | ");
                intNoOfItemsDeleted++;
            }

            Console.WriteLine("No of Entries Processed : {0} | Deleted : {1}", intNoOfItemsProcessed, intNoOfItemsDeleted);
            strDeleteResult = "Success";

            return strDeleteResult;
        }
        /// <summary>
        /// Load TM Lists.
        /// </summary>
        /// <param name="oClientContext">Client Context</param>
        /// <param name="strListName">List Name</param>
        /// <returns><c>true</c>If Dictionaries Loaded, otherwise <c>false</c></returns>
        private static string LoadTMLists(ClientContext oClientContext, string strListName)
        {
            string strLoadListResult = "false";
            Int32 intItemsLoaded = 0;
            ListItemCollectionPosition licp = null;            
            Web oWeb = oClientContext.Web;

            List oList = oWeb.Lists.GetByTitle(strListName);

            oClientContext.Load(oList);

            CamlQuery oQuery = new CamlQuery();
            oQuery.ViewXml = @"<View><ViewFields><FieldRef Name='Title'/><FieldRef Name='Nutrition_x0020_TM'/><FieldRef Name='Portfolio_x0020_TM'/><FieldRef Name='Pharmacy_x0020_TM'/></ViewFields></View>";
            oQuery.ListItemCollectionPosition = licp;
            ListItemCollection oItems = oList.GetItems(oQuery);

            oClientContext.Load(oItems);
            oClientContext.ExecuteQuery();

            string strPostalCode, strNutriTM, strPortTM, strPharmacyTM;

            dictNutriTM = new Dictionary<string, string>();
            dictPortTM = new Dictionary<string, string>();
            dictPharmacyTM = new Dictionary<string, string>();

            for (int i = 0; i < oItems.Count; i++)
            {
                strPostalCode = "";
                strNutriTM = "";
                strPortTM = "";
                strPharmacyTM = "";

                if(oItems[i]["Title"] != null)
                {
                    strPostalCode = oItems[i]["Title"].ToString();
                    if (dictNutriTM.ContainsKey(strPostalCode) == true)
                    {
                        Console.WriteLine(strPostalCode);
                    }
                    else
                    {
                        strNutriTM = (oItems[i]["Nutrition_x0020_TM"] == null) ? "" : oItems[i]["Nutrition_x0020_TM"].ToString();
                        strPortTM = (oItems[i]["Portfolio_x0020_TM"] == null) ? "" : oItems[i]["Portfolio_x0020_TM"].ToString();
                        strPharmacyTM = (oItems[i]["Pharmacy_x0020_TM"] == null) ? "" : oItems[i]["Pharmacy_x0020_TM"].ToString();

                        dictNutriTM.Add(strPostalCode, strNutriTM);
                        dictPortTM.Add(strPostalCode, strPortTM);
                        dictPharmacyTM.Add(strPostalCode, strPharmacyTM);

                        intItemsLoaded++;
                    }                    
                }

            }

            Console.WriteLine("[{0}]Items from List [{1}] loaded successfully.", intItemsLoaded, strListName);

            strLoadListResult = "true";

            return strLoadListResult;
        }

         /// <summary>
        /// Load TM Lists.
        /// </summary>
        /// <param name="oClientContext">Client Context</param>
        /// <param name="strListName">List Name</param>
        /// <returns><c>true</c>If Dictionaries Loaded, otherwise <c>false</c></returns>
        private static string GetTMValue(string strPostalCode, int TMIndexCode)
        {
            string strTMValue = "";

            
            switch(TMIndexCode)
            {
                case 1://Nutrition TM
                    if(dictNutriTM.ContainsKey(strPostalCode))
                    {
                        strTMValue = dictNutriTM[strPostalCode];
                    }
                    break;

                case 2://Portfolio TM
                    if (dictPortTM.ContainsKey(strPostalCode))
                    {
                        strTMValue = dictPortTM[strPostalCode];
                    }
                    break;

                case 3://Pharmacy TM
                    if (dictPharmacyTM.ContainsKey(strPostalCode))
                    {
                        strTMValue = dictPharmacyTM[strPostalCode];
                    }
                    break;
            }

            return strTMValue;
        }

        /// <summary>
        /// Check whether the Customer No exists in the Customer List.
        /// </summary>
        /// <param name="lstCustomerNo">To be Removed Customer No List</param>
        /// <param name="strCustomerNo">Customer No</param>
        /// <c>true</c>If Customer No exists in Remove List, otherwise <c>false</c></returns>
        private static bool IsItemExists(List<string> lstCustomerNo, string strCustomerNo)
        {
            bool doItemExists = false;

            foreach(string strCustomerNoTmp in lstCustomerNo)
            {
                if(strCustomerNoTmp == strCustomerNo)
                {
                    doItemExists = true;
                    break;
                }
            }

            return doItemExists;
        }

         /// <summary>
        /// Add New Item in Back Order List
        /// </summary>
        /// <param name="oClientContext">Client Context Object</param>
        /// <param name="oList">List Object</param>
        /// <param name="aryCellValues">Cell Values Array</param>
        /// <returns>Result</returns>
        public static Int64 AddNewBackOrderItem(ClientContext oClientContext, Web oWeb, string strListName, string strExcelFilePath)
        {
            Int64 intNoOfItemsUpdated = 0;
            //string strItemAddedResult = "Init\n";
           // try
            {
                List oList = oWeb.Lists.GetByTitle(strListName);
                CamlQuery caml = new CamlQuery();
                oClientContext.Load(oList);
                oClientContext.ExecuteQuery();

                UploadDataThread objUploadDataThread = new UploadDataThread();

                Thread[] thUploadDataThread = new Thread[5];

                //var intTempThreadCounter = 0;
                //const int NoOfThreadsLimit = 5;

                //ManualResetEvent[] doneEvents = new ManualResetEvent[NoOfThreadsLimit];
                //UploadDataThread[] aryUDThread = new UploadDataThread[NoOfThreadsLimit];

                Application ExcelObj = new Application();

                Workbook theWorkbook = ExcelObj.Workbooks.Open(strExcelFilePath, 0, true, 5,
                         "", "", true, XlPlatform.xlWindows, "\t", false, false,
                         0, true);
                Sheets sheets = theWorkbook.Worksheets;
                Worksheet worksheet = (Worksheet)sheets.get_Item(1);

                bool isMoreItemToAdd = true;
                Int32 intItemIndex = 5;


                while (isMoreItemToAdd == true)
                {
                    Range range = worksheet.get_Range("A" + intItemIndex.ToString(), "P" + intItemIndex.ToString());
                    System.Array aryCellValues = (System.Array)range.Cells.Value;
                    
                    if (aryCellValues.GetValue(1, 1) == null)
                    {
                        //if (intTempThreadCounter > 0)
                        //{
                        //    Console.WriteLine("Upload Thread Count at end " + intTempThreadCounter);
                        //    foreach (var e in doneEvents)
                        //    {
                        //        if (e != null)
                        //        {
                        //            e.WaitOne();
                        //        }
                        //    }
                        //}
                        isMoreItemToAdd = false;
                    }
                    else
                    {
                        Console.Write(AddNewBackOrderItem(oClientContext, oList, aryCellValues));
                        intNoOfItemsUpdated++;
                        //doneEvents[intTempThreadCounter] = new ManualResetEvent(false);
                        //UploadDataThread objUDT = new UploadDataThread(oClientContext, oList, worksheet, intItemIndex, doneEvents[intTempThreadCounter]);
                        //aryUDThread[intTempThreadCounter] = objUDT;
                        //ThreadPool.QueueUserWorkItem(objUDT.ThreadPoolCallback, intTempThreadCounter);

                        //intTempThreadCounter++;

                        //if (intTempThreadCounter > (NoOfThreadsLimit - 1))
                        //{
                        //    foreach (var e in doneEvents)
                        //    {
                        //        if (e != null)
                        //        {
                        //            e.WaitOne();
                        //        }
                        //    }
                        //    Console.WriteLine("Upload Thread Reset. " + intTempThreadCounter);
                        //    intTempThreadCounter = 0;
                        //}

                    }
                    intItemIndex++;
                }                

                theWorkbook.Close();

                Console.WriteLine("Upload Data Completed. Rows updated : ", (intItemIndex - 5));

            }
           //catch (Exception ex)
           // {
           //     strItemAddedResult = "Upload Data Failure : " + ex.Message;
           // }


            return intNoOfItemsUpdated;  
        }

        /// <summary>
        /// Add New Item in Allocations List
        /// </summary>
        /// <param name="oClientContext">Client Context Object</param>
        /// <param name="oList">List Object</param>
        /// <param name="aryCellValues">Cell Values Array</param>
        /// <returns>Result</returns>
        public static Int64 AddNewAllocationItem(ClientContext oClientContext, Web oWeb, string strListName, string strExcelFilePath)
        {
            Int64 intNoOfItemsUpdated = 0;
            //string strItemAddedResult = "Init\n";
            // try
            {
                List oList = oWeb.Lists.GetByTitle(strListName);
                CamlQuery caml = new CamlQuery();
                oClientContext.Load(oList);
                oClientContext.ExecuteQuery();

                UploadDataThread objUploadDataThread = new UploadDataThread();

                int intTempThreadCounter = 0, intThreadCount = 1;

                const int NoOfThreadsLimit = 10;
                Thread[] thUploadDataThread = new Thread[10];
                
                ManualResetEvent[] doneEvents = new ManualResetEvent[NoOfThreadsLimit];
                UploadDataThread[] aryUDThread = new UploadDataThread[NoOfThreadsLimit];

                Application ExcelObj = new Application();

                Workbook theWorkbook = ExcelObj.Workbooks.Open(strExcelFilePath, 0, true, 5,
                         "", "", true, XlPlatform.xlWindows, "\t", false, false,
                         0, true);
                Sheets sheets = theWorkbook.Worksheets;
                Worksheet worksheetWeekly = (Worksheet)sheets.get_Item(1);

                bool isMoreItemToAdd = true;
                Int32 intItemIndex = 2;
                string recurrence = "Weekly";

                while (isMoreItemToAdd == true)
                {
                    Range range = worksheetWeekly.get_Range("A" + intItemIndex.ToString(), "F" + intItemIndex.ToString());
                    System.Array aryCellValues = (System.Array)range.Cells.Value;

                    if (aryCellValues.GetValue(1, 1) == null)
                    {
                        if (intTempThreadCounter > 0)
                        {
                            //Console.WriteLine("Upload Thread Count at end " + intTempThreadCounter);
                            foreach (var e in doneEvents)
                            {
                                if (e != null)
                                {
                                    e.WaitOne(1500);
                                }
                            }
                        }
                        isMoreItemToAdd = false;
                    }
                    else
                    {
                        //Console.Write(AddNewBackOrderItem(oClientContext, oList, aryCellValues));
                        doneEvents[intTempThreadCounter] = new ManualResetEvent(false);
                        UploadDataThread objUDT = new UploadDataThread(oClientContext, oList, aryCellValues, intItemIndex, doneEvents[intTempThreadCounter], recurrence);
                        aryUDThread[intTempThreadCounter] = objUDT;
                        ThreadPool.QueueUserWorkItem(objUDT.ThreadPoolCallback, intTempThreadCounter);

                        intTempThreadCounter++;

                        if (intTempThreadCounter > (NoOfThreadsLimit - 1))
                        {
                            foreach (var e in doneEvents)
                            {
                                if (e != null)
                                {
                                    e.WaitOne(1500);
                                }
                            }
                            Console.WriteLine("Upload Thread Reset. " + intThreadCount);
                            intTempThreadCounter = 0;
                            intThreadCount++;
                        }
                        intNoOfItemsUpdated++;
                    }
                    intItemIndex++;
                }

                Console.WriteLine("Upload Weekly Data Completed. Rows updated : ", (intItemIndex - 1));
                                

                if (sheets.Count > 1)
                {
                    recurrence = "Monthly";

                    ManualResetEvent[] doneEventsMonthly = new ManualResetEvent[NoOfThreadsLimit];
                    UploadDataThread[] aryUDThreadMonthly = new UploadDataThread[NoOfThreadsLimit];

                    Worksheet worksheetMonthly = (Worksheet)sheets.get_Item(2);
                    intItemIndex = 2;
                    intTempThreadCounter = 0;
                    intThreadCount = 1;

                    isMoreItemToAdd = true;

                    while (isMoreItemToAdd == true)
                    {
                        Range range = worksheetMonthly.get_Range("A" + intItemIndex.ToString(), "F" + intItemIndex.ToString());
                        System.Array aryCellValues = (System.Array)range.Cells.Value;

                        if (aryCellValues.GetValue(1, 1) == null)
                        {
                            if (intTempThreadCounter > 0)
                            {
                                //Console.WriteLine("Upload Thread Count at end " + intTempThreadCounter);
                                foreach (var e in doneEventsMonthly)
                                {
                                    if (e != null)
                                    {
                                        e.WaitOne(1500);
                                    }
                                }
                            }
                            isMoreItemToAdd = false;
                        }
                        else
                        {
                            //Console.Write(AddNewBackOrderItem(oClientContext, oList, aryCellValues));
                            doneEventsMonthly[intTempThreadCounter] = new ManualResetEvent(false);
                            UploadDataThread objUDT = new UploadDataThread(oClientContext, oList, aryCellValues, intItemIndex, doneEventsMonthly[intTempThreadCounter], recurrence);
                            aryUDThreadMonthly[intTempThreadCounter] = objUDT;
                            ThreadPool.QueueUserWorkItem(objUDT.ThreadPoolCallback, intTempThreadCounter);

                            intTempThreadCounter++;

                            if (intTempThreadCounter > (NoOfThreadsLimit - 1))
                            {
                                foreach (var e in doneEventsMonthly)
                                {
                                    if (e != null)
                                    {
                                        e.WaitOne();
                                    }
                                }
                                Console.WriteLine("Upload Thread Reset. " + intThreadCount);
                                intTempThreadCounter = 0;
                                intThreadCount++;
                            }

                            intNoOfItemsUpdated++;
                        }
                        intItemIndex++;
                    }

                }

                theWorkbook.Close();                

                Console.WriteLine("Upload Monthly Data Completed. Rows updated : ", (intItemIndex - 1 ));

            }
            //catch (Exception ex)
            // {
            //     strItemAddedResult = "Upload Data Failure : " + ex.Message;
            // }


            return intNoOfItemsUpdated;
        }

        /// <summary>
        /// Add New Item in Allocations List
        /// </summary>
        /// <param name="oClientContext">Client Context Object</param>
        /// <param name="oList">List Object</param>
        /// <param name="aryCellValues">Cell Values Array</param>
        /// <returns>Result</returns>
        public static Int64 AddNewAllocationItemBCK(ClientContext oClientContext, Web oWeb, string strListName, string strExcelFilePath)
        {
            Int64 intNoOfItemsUpdated = 0;
            //string strItemAddedResult = "Init\n";
            // try
            {
                List oList = oWeb.Lists.GetByTitle(strListName);
                CamlQuery caml = new CamlQuery();
                oClientContext.Load(oList);
                oClientContext.ExecuteQuery();

                UploadDataThread objUploadDataThread = new UploadDataThread();

                int intTempThreadCounter = 0, intThreadCount = 1;

                const int NoOfThreadsLimit = 10;
                Thread[] thUploadDataThread = new Thread[10];

                ManualResetEvent[] doneEvents = new ManualResetEvent[NoOfThreadsLimit];
                UploadDataThread[] aryUDThread = new UploadDataThread[NoOfThreadsLimit];

                Application ExcelObj = new Application();

                Workbook theWorkbook = ExcelObj.Workbooks.Open(strExcelFilePath, 0, true, 5,
                         "", "", true, XlPlatform.xlWindows, "\t", false, false,
                         0, true);
                Sheets sheets = theWorkbook.Worksheets;
                Worksheet worksheetWeekly = (Worksheet)sheets.get_Item(1);

                bool isMoreItemToAdd = true;
                Int32 intItemIndex = 2;
                string recurrence = "Weekly";

                while (isMoreItemToAdd == true)
                {
                    Range range = worksheetWeekly.get_Range("A" + intItemIndex.ToString(), "F" + intItemIndex.ToString());
                    System.Array aryCellValues = (System.Array)range.Cells.Value;

                    if (aryCellValues.GetValue(1, 1) == null)
                    {
                        if (intTempThreadCounter > 0)
                        {
                            //Console.WriteLine("Upload Thread Count at end " + intTempThreadCounter);
                            foreach (var e in doneEvents)
                            {
                                if (e != null)
                                {
                                    e.WaitOne();
                                }
                            }
                        }
                        isMoreItemToAdd = false;
                    }
                    else
                    {
                        //Console.Write(AddNewBackOrderItem(oClientContext, oList, aryCellValues));
                        doneEvents[intTempThreadCounter] = new ManualResetEvent(false);
                        UploadDataThread objUDT = new UploadDataThread(oClientContext, oList, aryCellValues, intItemIndex, doneEvents[intTempThreadCounter], recurrence);
                        aryUDThread[intTempThreadCounter] = objUDT;
                        ThreadPool.QueueUserWorkItem(objUDT.ThreadPoolCallback, intTempThreadCounter);

                        intTempThreadCounter++;

                        if (intTempThreadCounter > (NoOfThreadsLimit - 1))
                        {
                            foreach (var e in doneEvents)
                            {
                                if (e != null)
                                {
                                    e.WaitOne();
                                }
                            }
                            Console.WriteLine("Upload Thread Reset. " + intThreadCount);
                            intTempThreadCounter = 0;
                            intThreadCount++;
                        }
                        intNoOfItemsUpdated++;
                    }
                    intItemIndex++;
                }

                Console.WriteLine("Upload Weekly Data Completed. Rows updated : ", (intItemIndex - 1));


                if (sheets.Count > 1)
                {
                    recurrence = "Monthly";

                    ManualResetEvent[] doneEventsMonthly = new ManualResetEvent[NoOfThreadsLimit];
                    UploadDataThread[] aryUDThreadMonthly = new UploadDataThread[NoOfThreadsLimit];

                    Worksheet worksheetMonthly = (Worksheet)sheets.get_Item(2);
                    intItemIndex = 2;
                    intTempThreadCounter = 0;
                    intThreadCount = 1;

                    isMoreItemToAdd = true;

                    while (isMoreItemToAdd == true)
                    {
                        Range range = worksheetMonthly.get_Range("A" + intItemIndex.ToString(), "F" + intItemIndex.ToString());
                        System.Array aryCellValues = (System.Array)range.Cells.Value;

                        if (aryCellValues.GetValue(1, 1) == null)
                        {
                            if (intTempThreadCounter > 0)
                            {
                                //Console.WriteLine("Upload Thread Count at end " + intTempThreadCounter);
                                foreach (var e in doneEventsMonthly)
                                {
                                    if (e != null)
                                    {
                                        e.WaitOne();
                                    }
                                }
                            }
                            isMoreItemToAdd = false;
                        }
                        else
                        {
                            //Console.Write(AddNewBackOrderItem(oClientContext, oList, aryCellValues));
                            doneEventsMonthly[intTempThreadCounter] = new ManualResetEvent(false);
                            UploadDataThread objUDT = new UploadDataThread(oClientContext, oList, aryCellValues, intItemIndex, doneEventsMonthly[intTempThreadCounter], recurrence);
                            aryUDThreadMonthly[intTempThreadCounter] = objUDT;
                            ThreadPool.QueueUserWorkItem(objUDT.ThreadPoolCallback, intTempThreadCounter);

                            intTempThreadCounter++;

                            if (intTempThreadCounter > (NoOfThreadsLimit - 1))
                            {
                                foreach (var e in doneEventsMonthly)
                                {
                                    if (e != null)
                                    {
                                        e.WaitOne();
                                    }
                                }
                                Console.WriteLine("Upload Thread Reset. " + intThreadCount);
                                intTempThreadCounter = 0;
                                intThreadCount++;
                            }

                            intNoOfItemsUpdated++;
                        }
                        intItemIndex++;
                    }

                }

                theWorkbook.Close();

                Console.WriteLine("Upload Monthly Data Completed. Rows updated : ", (intItemIndex - 1));

            }
            //catch (Exception ex)
            // {
            //     strItemAddedResult = "Upload Data Failure : " + ex.Message;
            // }


            return intNoOfItemsUpdated;
        }


        /// <summary>
        /// Concatenate all the elements into a StringBuilder.
        /// </summary>
        /// <param name="aryCellValues">Cell Values Array</param>
        /// <param name="index">Index</param>
        /// <returns>Concatenated values</returns>
        private static string ConvertToStringArray(Array aryCellValues, int index)
        {
            
            StringBuilder builder = new StringBuilder();
            for (int i = 1; i < aryCellValues.Length; i++)
            {
                builder.Append(Convert.ToString(aryCellValues.GetValue(index,i)));
                builder.Append(" | ");
            }  
            return builder.ToString();
        }
        
        /// <summary>
        /// Add New Item in Back Order List
        /// </summary>
        /// <param name="oClientContext">Client Context Object</param>
        /// <param name="oList">List Object</param>
        /// <param name="aryCellValues">Cell Values Array</param>
        /// <returns>Result</returns>
        public static string AddNewBackOrderItem(ClientContext oClientContext, List oList, Array aryCellValues)
        {
            string strItemAddedResult = "Init\n";
            string[] aryBackOrderFields = "Title|Account_x0020_Name|City|State|Postal_x0020_Code|Customer_x0020_Type|Customer_x0020_Type_x0020_Descri|Order_x0020_Date|Order_x0020_No|PO_x0020_No|Release_x0020_No|Item_x0020_No|Description|Qty_x0020_Ordered|Qty_x0020_BackOrdered|Warehouse".Split('|');
            try
            {
                string strPostalCode = "";
                ListItemCreationInformation itmCreateInfo = new ListItemCreationInformation();
                ListItem liNewItem = oList.AddItem(itmCreateInfo);
                for (int i = 0; i < aryBackOrderFields.Length; i++)
                {
                    liNewItem[aryBackOrderFields[i]] = (aryCellValues.GetValue(1, i + 1) != null)? aryCellValues.GetValue(1, i + 1).ToString().Trim(): "";

                    if(i == 4)
                    {
                        strPostalCode = (aryCellValues.GetValue(1, i + 1) != null) ? aryCellValues.GetValue(1, i + 1).ToString().Trim() : "";
                    }
                }

                // Update TM Sales Rep Values
                if (strPostalCode != "")
                {
                    liNewItem["Nutrition_x0020_TM"] = GetTMValue(strPostalCode, 1);
                    liNewItem["Portfolio_x0020_TM"] = GetTMValue(strPostalCode, 2);
                    liNewItem["Pharmacy_x0020_TM"] = GetTMValue(strPostalCode, 3);
                }

                liNewItem.Update();

                oClientContext.ExecuteQuery();
                intItemCount++;

                strItemAddedResult = ".";//Success.
            }
            catch(Exception ex)
            {
                strItemAddedResult = "Failure : " + ex.Message;
            }
            

            return strItemAddedResult;    
        }
    }

    /// <summary>
    /// Delete Data on Threads Class
    /// </summary>

    public class DeleteDataThread
    {
        private ManualResetEvent _doneEvent;

        public ListItemCollection _oItems;
        public ClientContext _oClientContext;


        public DeleteDataThread()
        {
        }

        // Constructor. 
        public DeleteDataThread(ClientContext oClientContext, ListItemCollection oItems, ManualResetEvent doneEvent)
        {
            _oClientContext = oClientContext;
            _oItems = oItems;
            _doneEvent = doneEvent;
        }

        // Wrapper method for use with thread pool. 
        public void ThreadPoolCallback(Object threadContext)
        {
            int threadIndex = (int)threadContext;
            if (_oItems != null && _oItems.Count > 0)
            {
                DeleteListItems(_oItems, _oClientContext);
            }
            _doneEvent.Set();
        }


        /// <summary>
        /// Delete List Items from SP List
        /// </summary>
        /// <param name="oItems">List Items Collection</param>
        /// <param name="oClientContext">SharePoint Collection</param>
        /// <returns><c>true</c>If items deleted, otherwise <c>false</c></returns>
        public bool DeleteListItems(ListItemCollection oItems, ClientContext oClientContext)
        {
            bool isItemDeleted = false;
            foreach (ListItem oItm in oItems.ToList())
            {
                oItm.DeleteObject();
                isItemDeleted = true;
            }
            oClientContext.ExecuteQuery();
            Console.Write('-');
            return isItemDeleted;
        }        
    }


    /// <summary>
    /// Upload Data on Threads Class
    /// </summary>

    public class UploadDataThread
    {
        string[] aryBackOrderFields, aryAllocationFields;
        private ManualResetEvent _doneEvent;

        public List _oList;
        public ClientContext _oClientContext;
        public Worksheet _worksheet;
        public System.Array _aryCellValues;
        public Int32 _intItemIndex;
        public string _recurrence;

        public UploadDataThread()
        {
            //aryBackOrderFields = "Title|Account_x0020_Name|City|State|Customer_x0020_Type|Customer_x0020_Type_x0020_Descri|Order_x0020_Date|Order_x0020_No|PO_x0020_No|Release_x0020_No|Item_x0020_No|Description|Qty_x0020_Ordered|Qty_x0020_BackOrdered|Warehouse".Split('|');
        }

        // Constructor. 
        public UploadDataThread(ClientContext oClientContext, List oList, Worksheet worksheet, Int32 intItemIndex, ManualResetEvent doneEvent)
        {            
            aryAllocationFields = "Title|Customer_x0020_Group|Item_x0020_No|Quantity_x0020_Limit|Quantity_x0020_Sold|Expiration_x0020_Date|Recurrence".Split('|');
            _oClientContext = oClientContext;
            _oList = oList;
            _worksheet = worksheet;
            _intItemIndex = intItemIndex;
            _doneEvent = doneEvent;
        }

        // Constructor. 
        public UploadDataThread(ClientContext oClientContext, List oList, System.Array aryCellValues, Int32 intItemIndex, ManualResetEvent doneEvent, string recurrence)
        {            
            aryAllocationFields = "Title|Customer_x0020_Group|Item_x0020_No|Quantity_x0020_Limit|Quantity_x0020_Sold|Expiration_x0020_Date|Recurrence".Split('|');
            _oClientContext = oClientContext;
            _oList = oList;
            _aryCellValues = aryCellValues;
            _recurrence = recurrence;
            _intItemIndex = intItemIndex;
            _doneEvent = doneEvent;
        }

        // Wrapper method for use with thread pool. 
        public void ThreadPoolCallback(Object threadContext)
        {
            int threadIndex = (int)threadContext;
            FetchValueAndUploadAllocations(_aryCellValues, _oClientContext, _oList, _recurrence, ref _doneEvent);            
        }

        /// <summary>
        /// Fetch values from Excel and Upload on SP List
        /// </summary>
        /// <param name="worksheet">Excel Worksheet Reference</param>
        /// <param name="intItemIndex">Item Index</param>
        /// <param name="oClientContext">SharePoint Client Context</param>
        /// <param name="oList">SharePoint List</param>
        /// <returns><c>true</c>If still add more items, otherwise <c>false</c></returns>
        public bool FetchValueAndUploadAllocations(System.Array aryCellValues, ClientContext oClientContext, List oList, string recurrence, ref ManualResetEvent doneEvent)
        {
            bool isMoreItemToAdd = true;
            
            if (aryCellValues.GetValue(1, 1) == null)
            {
                isMoreItemToAdd = false;            
            }
            else
            {
                Console.Write(AddNewAllocationItem(oClientContext, oList, aryCellValues, recurrence));                
            }
            doneEvent.Set();

            return isMoreItemToAdd;
        }

        /// <summary>
        /// Add New Item in Back Order List
        /// </summary>
        /// <param name="oClientContext">Client Context Object</param>
        /// <param name="oList">List Object</param>
        /// <param name="aryCellValues">Cell Values Array</param>
        /// <returns>Result</returns>
        private string AddNewBackOrderItem(ClientContext oClientContext, List oList, Array aryCellValues)
        {
            string strItemAddedResult = "Init\n";

            //try
            {
                ListItemCreationInformation itmCreateInfo = new ListItemCreationInformation();
                ListItem liNewItem = oList.AddItem(itmCreateInfo);
                for (int i = 0; i < aryBackOrderFields.Length; i++)
                {
                    liNewItem[aryBackOrderFields[i]] = aryCellValues.GetValue(1, i + 1);
                }

                liNewItem.Update();

                oClientContext.ExecuteQuery();

                strItemAddedResult = ".";//Success.
            }
            //catch (Exception ex)
            //{
            //    strItemAddedResult = "Failure : " + ex.Message;
            //}

            return strItemAddedResult;
        }

        /// <summary>
        /// Add New Item in Allocations List
        /// </summary>
        /// <param name="oClientContext">Client Context Object</param>
        /// <param name="oList">List Object</param>
        /// <param name="aryCellValues">Cell Values Array</param>
        /// <returns>Result</returns>
        private string AddNewAllocationItem(ClientContext oClientContext, List oList, Array aryCellValues, string recurrence)
        {
            string strItemAddedResult = "Init\n";

            if (aryCellValues.GetValue(1, 1) != null)
            {
                //try
                {
                    ListItemCreationInformation itmCreateInfo = new ListItemCreationInformation();
                    ListItem liNewItem = oList.AddItem(itmCreateInfo);
                    for (int i = 0; i < aryAllocationFields.Length - 1; i++)
                    {
                        liNewItem[aryAllocationFields[i]] = aryCellValues.GetValue(1, i + 1);
                    }

                    //Update Recurrence [Weekly/Monthly]
                    liNewItem[aryAllocationFields[aryAllocationFields.Length - 1]] = recurrence;

                    liNewItem.Update();

                    oClientContext.ExecuteQuery();

                    strItemAddedResult = "-";//Success.
                }
                //catch (Exception ex)
                //{
                //    strItemAddedResult = "\n " + aryCellValues.GetValue(1, 1) + " : " + ex.Message;
                //}
            }

            return strItemAddedResult;
        }
    }
}

#region Unused Code
//FieldCollection collFields = oList.Fields;                
/*oClientContext.LoadQuery(List => List.Fields.Include(
        field => field.Title,
        field => field.InternalName));
*/
/*
foreach (Field f in oList.Fields)
{
    Console.WriteLine("{0} - {1}", f.Title, f.InternalName, f.Hidden, f.CanBeDeleted);
}
               
foreach (Field oField in oList.Fields)
{
    Regex regEx = new Regex("name", RegexOptions.IgnoreCase);
                    
    if (regEx.IsMatch(oField.InternalName))
    {
        Console.WriteLine("Field Title: {0} \n\t Field Internal Name: {1}", 
            oField.Title, oField.InternalName);
    }
}
*/
/*
ListItemCollection items = oList.GetItems(caml);

cc.Load<List>(oList);
cc.Load<ListItemCollection>(items);
cc.ExecuteQuery();

foreach (Microsoft.SharePoint.Client.ListItem item in items)
{
    Console.WriteLine(item.FieldValues["Title"]);
}

public static string AddNewBackOrderItem(ClientContext oClientContext, Web oWeb, string strListName, string strExcelFilePath)
        {
            string strItemAddedResult = "Init\n";
           // try
            {
                List oList = oWeb.Lists.GetByTitle(strListName);
                CamlQuery caml = new CamlQuery();
                oClientContext.Load(oList);
                oClientContext.ExecuteQuery();

                UploadDataThread objUploadDataThread = new UploadDataThread();

                Thread[] thUploadDataThread = new Thread[5];

                var intTempThreadCounter = 0;
                const int NoOfThreadsLimit = 5;

                ManualResetEvent[] doneEvents = new ManualResetEvent[NoOfThreadsLimit];
                UploadDataThread[] aryUDThreadMonthly = new UploadDataThread[NoOfThreadsLimit];

                Application ExcelObj = new Application();

                Workbook theWorkbook = ExcelObj.Workbooks.Open(strExcelFilePath, 0, true, 5,
                         "", "", true, XlPlatform.xlWindows, "\t", false, false,
                         0, true);
                Sheets sheets = theWorkbook.Worksheets;
                Worksheet worksheet = (Worksheet)sheets.get_Item(1);

                bool isMoreItemToAdd = true;
                Int32 intItemIndex = 5;


                while (isMoreItemToAdd == true)
                {
                    Range range = worksheet.get_Range("A" + intItemIndex.ToString(), "O" + intItemIndex.ToString());
                    System.Array aryCellValues = (System.Array)range.Cells.Value;

                    if (aryCellValues.GetValue(1, 1) == null)
                    {
                        if (intTempThreadCounter > 0)
                        {
                            Console.WriteLine("Upload Thread Count at end " + intTempThreadCounter);
                            foreach (var e in doneEvents)
                            {
                                if (e != null)
                                {
                                    e.WaitOne();
                                }
                            }
                        }
                        isMoreItemToAdd = false;
                    }
                    else
                    {
                        doneEvents[intTempThreadCounter] = new ManualResetEvent(false);
                        UploadDataThread objUDT = new UploadDataThread(oClientContext, oList, worksheet, intItemIndex, doneEvents[intTempThreadCounter]);
                        aryUDThreadMonthly[intTempThreadCounter] = objUDT;
                        ThreadPool.QueueUserWorkItem(objUDT.ThreadPoolCallback, intTempThreadCounter);

                        intTempThreadCounter++;

                        if (intTempThreadCounter > (NoOfThreadsLimit - 1))
                        {
                            foreach (var e in doneEvents)
                            {
                                if (e != null)
                                {
                                    e.WaitOne();
                                }
                            }
                            Console.WriteLine("Upload Thread Reset. " + intTempThreadCounter);
                            intTempThreadCounter = 0;
                        }

                    }
                    intItemIndex++;
                }                

                theWorkbook.Close();

                Console.WriteLine("Upload Data Completed.");

            }
           //catch (Exception ex)
           // {
           //     strItemAddedResult = "Upload Data Failure : " + ex.Message;
           // }


            return strItemAddedResult;  
        }
 * */
#endregion
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;

namespace SharePoint2010Snippets_Console
{
    public class FolderDetails
    {
        public FolderDetails()
        {
        }

        //Method - Get SubFolder Count from Parent Fodler
        public void getfoldercount(string folderurl)
        {
            try
            {
                using (SPSite site = new SPSite(folderurl))
                {
                    using (SPWeb web = site.OpenWeb())
                    {
                    
                    SPFolder folder = web.GetFolder(folderurl);
                    Console.WriteLine("Folder: " + folder.Name);
                    //Get Sub-Folders Count

                    Console.WriteLine("Sub-Folders: " + folder.Properties["vti_foldersubfolderitemcount"].ToString());
                    //Get Sub-Files Count
                    //Console.WriteLine(folder.Files.Count.ToString());
                    }
                }                
            }
            catch (Exception ex)
            {
                Console.WriteLine("Snippet failed! Error: " + ex.Message);               
            }
            Console.WriteLine("Press Enter to Continue or Exit...");
            Console.Read();
        }
    }
}

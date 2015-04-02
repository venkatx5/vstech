using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;

namespace SharePoint2010Snippets_Console
{
    public class FileDetails
    {

        public FileDetails()
        {
        }

        //Check the file is exists in SharePoint
        public  void isfileexists(string fileurl)
        {
            try
            {
                using (SPSite site = new SPSite("http://localhost"))
                {
                    using (SPWeb web = site.OpenWeb())
                    {
                        SPFile fle = web.GetFile("/SitePages/Home.aspx");

                        if (fle.Exists)
                            Console.WriteLine("File \"" + site.Url +  fle.ServerRelativeUrl + "\" Exits");
                        else
                            Console.WriteLine(fle.ServerRelativeUrl + "No Exits");
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

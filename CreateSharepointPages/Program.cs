using Microsoft.SharePoint;
using Microsoft.SharePoint.Publishing;
using Microsoft.SharePoint.WebPartPages;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.UI.WebControls.WebParts;
using System.Xml;

namespace CreateSharepointPages
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("1.Create Pages (Default)");
            Console.WriteLine("2.Delete Pages");
            string strVal= Console.ReadLine();
            Console.WriteLine("Please enter site url");
            string strSiteUrl = Console.ReadLine();
            if (strVal == "2")
            {
                DeleteAllPages(strSiteUrl);
            }
            else
            {
                CreateAllPages(strSiteUrl);
            }
            
        }
        private static void DeleteAllPages(string SiteUrl)
        {
            using (SPSite root = new SPSite(SiteUrl) )
            {
                using (SPWeb site = root.OpenWeb())
                {
                    deletesitepages(site, Constants.ConstantsHelper.ADD_USER_PAGE);
                    deletesitepages(site, Constants.ConstantsHelper.ASTMR_PAGE);
                    deletesitepages(site, Constants.ConstantsHelper.CHANGE_PASSWORD_PAGE);
                    deletesitepages(site, Constants.ConstantsHelper.COMPANY_LIST_PAGE);
                    deletesitepages(site, Constants.ConstantsHelper.DISPLAY_REPORTS_PAGE);
                    deletesitepages(site, Constants.ConstantsHelper.EDIT_COMPANY_PAGE);
                    deletesitepages(site, Constants.ConstantsHelper.EDIT_LAB_PAGE);
                    deletesitepages(site, Constants.ConstantsHelper.EDIT_PROGRAM_PAGE);
                    deletesitepages(site, Constants.ConstantsHelper.EDIT_USER_PAGE);
                    deletesitepages(site, Constants.ConstantsHelper.LAB_LIST_PAGE);
                    deletesitepages(site, Constants.ConstantsHelper.MANAGE_PROGRAM_PAGE);
                    deletesitepages(site, Constants.ConstantsHelper.MANAGE_TEST_PAGE);
                    deletesitepages(site, Constants.ConstantsHelper.MYPROFILE_PAGE);
                    deletesitepages(site, Constants.ConstantsHelper.PROGRAM_RESULT_UPLOAD_PAGE);
                    deletesitepages(site, Constants.ConstantsHelper.PROGRAM_LIST_PAGE);
                    deletesitepages(site, Constants.ConstantsHelper.RUN_REPORT_PAGE);
                    deletesitepages(site, Constants.ConstantsHelper.USER_LIST_PAGE);
                }

            }
            Console.WriteLine();
            Console.WriteLine();
            Console.Write("Press any key to continue");
            Console.Read();
        }
        private static void CreateAllPages(string SiteUrl)
        {
            
            using (SPSite root = new SPSite(SiteUrl) )
            {
                using (SPWeb site = root.OpenWeb())
                {

                    createpages(site, CreateSharepointPages.Constants.ConstantsHelper.ADD_USER_PAGE, CreateSharepointPages.Constants.ConstantsHelper.ADD_USER_WEBPART);
                    createpages(site, CreateSharepointPages.Constants.ConstantsHelper.ASTMR_PAGE, CreateSharepointPages.Constants.ConstantsHelper.REPRODUCIBILITY_WEBPART);
                    createpages(site, CreateSharepointPages.Constants.ConstantsHelper.CHANGE_PASSWORD_PAGE, CreateSharepointPages.Constants.ConstantsHelper.CHANGE_PASSWORD_WEBPART);
                    createpages(site, CreateSharepointPages.Constants.ConstantsHelper.COMPANY_LIST_PAGE, CreateSharepointPages.Constants.ConstantsHelper.COMPANY_LIST_WEBPART);
                    createpages(site, CreateSharepointPages.Constants.ConstantsHelper.DISPLAY_REPORTS_PAGE, CreateSharepointPages.Constants.ConstantsHelper.DISPLAY_REPORTS_WEBPART);
                    createpages(site, CreateSharepointPages.Constants.ConstantsHelper.EDIT_COMPANY_PAGE, CreateSharepointPages.Constants.ConstantsHelper.EDIT_COMPANY_WEBPART);
                    createpages(site, CreateSharepointPages.Constants.ConstantsHelper.EDIT_LAB_PAGE, CreateSharepointPages.Constants.ConstantsHelper.EDIT_LAB_WEBPART);
                    createpages(site, CreateSharepointPages.Constants.ConstantsHelper.EDIT_PROGRAM_PAGE, CreateSharepointPages.Constants.ConstantsHelper.EDIT_PROGRAM_WEBPART);
                    createpages(site, CreateSharepointPages.Constants.ConstantsHelper.EDIT_USER_PAGE, CreateSharepointPages.Constants.ConstantsHelper.EDIT_USER_WEBPART);
                    createpages(site, CreateSharepointPages.Constants.ConstantsHelper.LAB_LIST_PAGE, CreateSharepointPages.Constants.ConstantsHelper.LAB_LIST_WEBPART);
                    createpages(site, CreateSharepointPages.Constants.ConstantsHelper.MANAGE_PROGRAM_PAGE, CreateSharepointPages.Constants.ConstantsHelper.MANAGE_PROGRAM_WEBPART);
                    createpages(site, CreateSharepointPages.Constants.ConstantsHelper.MANAGE_TEST_PAGE, CreateSharepointPages.Constants.ConstantsHelper.MANAGE_TEST_WEBPART);
                    createpages(site, CreateSharepointPages.Constants.ConstantsHelper.MYPROFILE_PAGE, CreateSharepointPages.Constants.ConstantsHelper.MYPROFILE_WEBPART);
                    createpages(site, CreateSharepointPages.Constants.ConstantsHelper.PROGRAM_RESULT_UPLOAD_PAGE, CreateSharepointPages.Constants.ConstantsHelper.PROGRAM_RESULT_UPLOAD_WEBPART);
                    createpages(site, CreateSharepointPages.Constants.ConstantsHelper.PROGRAM_LIST_PAGE, CreateSharepointPages.Constants.ConstantsHelper.PROGRAM_LIST_WEBPART);
                    createpages(site, CreateSharepointPages.Constants.ConstantsHelper.RUN_REPORT_PAGE, CreateSharepointPages.Constants.ConstantsHelper.RUN_REPORT_WEBPART);
                    createpages(site, CreateSharepointPages.Constants.ConstantsHelper.USER_LIST_PAGE, CreateSharepointPages.Constants.ConstantsHelper.USER_LIST_WEBPART);
                }
            }
            Console.WriteLine();
            Console.WriteLine();
            Console.Write("Press any key to continue");
            Console.Read();
        }



        protected static void createpages(SPWeb site, string pagename, string webpartname)
        {
            
            site.AllowUnsafeUpdates = true;
            PublishingWeb publishingWeb = PublishingWeb.GetPublishingWeb(site);
            PublishingPageCollection pages = publishingWeb.GetPublishingPages();
            if (pages[publishingWeb.Url + "/Pages/" + pagename] == null)
            {
                Console.WriteLine("Creating " + pagename.Replace("_", " ")); 
                
                PageLayout selectedLayout = null;
                PageLayout[] layouts = publishingWeb.GetAvailablePageLayouts();
                foreach (PageLayout layout in layouts)
                {
                    if (layout.Name.Trim() == Constants.ConstantsHelper.PAGE_TEMPLATE)
                    {
                        selectedLayout = layout;
                    }

                }

                PublishingPage newPage = pages.Add(pagename, selectedLayout);
                newPage.ListItem.Update();
                newPage.Update();

                AddWebPartToPage(site, pagename, webpartname);
                newPage.CheckIn(webpartname + "is added");
                Console.WriteLine("Created " + pagename.Replace("_", " "));
            }
            else 
            {
                Console.WriteLine("Already exist " + pagename.Replace("_", " "));
            }

            site.AllowUnsafeUpdates = false;
             
        }

        public static void AddWebPartToPage(SPWeb web, string pageUrl, string webPartName)
        {
            using (SPLimitedWebPartManager webPartManager = web.GetLimitedWebPartManager("Pages/" + pageUrl, PersonalizationScope.Shared))
            {
                using (System.Web.UI.WebControls.WebParts.WebPart webPart = CreateWebPart(web, webPartName, webPartManager))
                {
                    if (!string.IsNullOrEmpty(webPart.ToString()))
                    {
                        Console.WriteLine("adding webpart " + webPartName.Replace("_", " ")); 
                        webPart.ChromeType = PartChromeType.None;
                        webPartManager.AddWebPart(webPart, "FullPage", 0);
                        Console.WriteLine("added webpart " + webPartName.Replace("_", " ")); 
                    }

                    // return webPart.ID;
                }

            }
        }

        public static System.Web.UI.WebControls.WebParts.WebPart CreateWebPart(SPWeb web, string webPartName, SPLimitedWebPartManager webPartManager)
        {
            SPQuery qry = new SPQuery();
            qry.Query = String.Format(CultureInfo.CurrentCulture,
                "<Where><Eq><FieldRef Name='FileLeafRef'/><Value Type='File'>{0}</Value></Eq></Where>", webPartName);

            SPList webPartGallery = null;

            if (null == web.ParentWeb)
            {
                webPartGallery = web.GetCatalog(SPListTemplateType.WebPartCatalog);
            }
            else
            {
                webPartGallery = web.Site.RootWeb.GetCatalog(SPListTemplateType.WebPartCatalog);
            }

            SPListItemCollection webParts = webPartGallery.GetItems(qry);
            XmlReader xmlReader = new XmlTextReader(webParts[0].File.OpenBinaryStream());
            string errorMsg;
            System.Web.UI.WebControls.WebParts.WebPart webPart = webPartManager.ImportWebPart(xmlReader, out errorMsg);
            return webPart;
        }

        protected static void deletesitepages(SPWeb site, string page)
        {
            
            site.AllowUnsafeUpdates = true;
            PublishingWeb publishingWeb = null;
            if (PublishingWeb.IsPublishingWeb(site))
            {
                publishingWeb = PublishingWeb.GetPublishingWeb(site);

                // get the pages collection
                PublishingPageCollection pages = publishingWeb.GetPublishingPages();

                if (pages[publishingWeb.Url + "/Pages/" + page] != null)
                {
                    // page exists so delete it
                    Console.WriteLine("Deleting " + page.Replace("_", " "));
                    PublishingPage currentpage = pages[publishingWeb.Url + "/Pages/" + page];
                    currentpage.ListItem.Delete();
                    publishingWeb.Update();
                    Console.WriteLine("Deleted " + page.Replace("_", " "));

                }
                else
                {
                    Console.WriteLine("Does not exist " + page.Replace("_", " "));
                }

            }
            site.AllowUnsafeUpdates = false;
            
        }
        
    }


}

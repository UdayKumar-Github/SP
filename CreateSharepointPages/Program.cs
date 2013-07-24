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
            Console.WriteLine();
            Console.WriteLine();
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
                    deletesitepages(site, Constants.ConstantsHelper.ADD_CONTROL_PAGE);                    
                    deletesitepages(site, Constants.ConstantsHelper.ADD_GROUP_CONTRACT_PAGE);
                    deletesitepages(site, Constants.ConstantsHelper.ADD_MANAGERIAL_REPORT_PAGE);
                    deletesitepages(site, Constants.ConstantsHelper.ADD_REPORT_VARIABLE_PAGE);
                    deletesitepages(site, Constants.ConstantsHelper.ADD_USER_PAGE);
                    deletesitepages(site, Constants.ConstantsHelper.ASTMR_PAGE);
                    deletesitepages(site, Constants.ConstantsHelper.CHANGE_PASSWORD_PAGE);
                    deletesitepages(site, Constants.ConstantsHelper.COMPANY_LIST_PAGE);
                    deletesitepages(site, Constants.ConstantsHelper.COPY_FIELD_PAGE);
                    deletesitepages(site, Constants.ConstantsHelper.COPY_METHOD_PAGE);
                    deletesitepages(site, Constants.ConstantsHelper.COPY_TEST_PAGE);
                    deletesitepages(site, Constants.ConstantsHelper.DATA_COLLECTION_DATA_ENTRY_PAGE);
                    deletesitepages(site, Constants.ConstantsHelper.DATA_COLLECTION_ENTRY_PAGE);
                    deletesitepages(site, Constants.ConstantsHelper.DATA_COLLECTION_RESULT_PAGE);
                    deletesitepages(site, Constants.ConstantsHelper.DATA_COLLECTION_SUBMIT_PAGE);                   
                    deletesitepages(site, Constants.ConstantsHelper.DISPLAY_REPORTS_PAGE);
                    deletesitepages(site, Constants.ConstantsHelper.DOWNLOAD_PDF_PAGE);                    
                    deletesitepages(site, Constants.ConstantsHelper.EDIT_COMPANY_PAGE);
                    deletesitepages(site, Constants.ConstantsHelper.EDIT_CONTRACT_PAGE);
                    deletesitepages(site, Constants.ConstantsHelper.EDIT_LAB_PAGE);
                    deletesitepages(site, Constants.ConstantsHelper.EDIT_MANAGERIAL_REPORT_PAGE);
                    deletesitepages(site, Constants.ConstantsHelper.EDIT_PROGRAM_PAGE);
                    deletesitepages(site, Constants.ConstantsHelper.EDIT_REPORT_VARIABLE_PAGE);
                    deletesitepages(site, Constants.ConstantsHelper.EDIT_USER_PAGE);
                    deletesitepages(site, Constants.ConstantsHelper.LAB_LIST_PAGE);
                    deletesitepages(site, Constants.ConstantsHelper.LATE_REGISTRANTS_PAGE);
                    deletesitepages(site, Constants.ConstantsHelper.LATEREGISTRANTS_COMMENT_PAGE);
                    deletesitepages(site, Constants.ConstantsHelper.MANAGE_FIELD_PAGE);
                    deletesitepages(site, Constants.ConstantsHelper.MANAGE_GROUPCONTRACT_PAGE);
                    deletesitepages(site, Constants.ConstantsHelper.MANAGE_MANAGERIAL_REPORT_PAGE);
                    deletesitepages(site, Constants.ConstantsHelper.MANAGE_METHOD_PAGE);
                    deletesitepages(site, Constants.ConstantsHelper.MANAGE_PROGRAM_PAGE);
                    deletesitepages(site, Constants.ConstantsHelper.MANAGE_RENEWALS_PAGE);
                    deletesitepages(site, Constants.ConstantsHelper.MANAGE_TEST_PAGE);
                    deletesitepages(site, Constants.ConstantsHelper.MYPROFILE_PAGE);
                    deletesitepages(site, Constants.ConstantsHelper.NOTIFICATION_HISTORY_PAGE);
                    deletesitepages(site, Constants.ConstantsHelper.PROGRAM_RESULT_UPLOAD_PAGE);
                    deletesitepages(site, Constants.ConstantsHelper.PROGRAM_LIST_PAGE);
                    deletesitepages(site, Constants.ConstantsHelper.REMOTE_TEMPLATE_VARIABLE_PAGE);                    
                    deletesitepages(site, Constants.ConstantsHelper.RUN_REPORT_PAGE);
                    deletesitepages(site, Constants.ConstantsHelper.SEND_TEST_RENEWALNOTICE_PAGE);
                    deletesitepages(site, Constants.ConstantsHelper.UNSUBMITTED_TEST_RESULTS_PAGE);                  
                    deletesitepages(site, Constants.ConstantsHelper.USER_LIST_PAGE);
                    deletesitepages(site, Constants.ConstantsHelper.VERSION_DIALOG_PAGE);                    
                    deletesitepages(site, Constants.ConstantsHelper.VIEW_MANAGERIAL_REPORT_PAGE);
                    deletesitepages(site, Constants.ConstantsHelper.VIEW_REPORT_PAGE);
                    deletesitepages(site, Constants.ConstantsHelper.VIEW_RESULTS_PAGE);


                    
                   
                    
                    
                   
                    
                    
                    

                    
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

                    createpages(site, Constants.ConstantsHelper.ADD_CONTROL_PAGE, Constants.ConstantsHelper.ADD_CONTROL_WEBPART);
                    createpages(site, Constants.ConstantsHelper.ADD_GROUP_CONTRACT_PAGE, Constants.ConstantsHelper.ADD_GROUP_CONTROL_WEBPART);
                    createpages(site, Constants.ConstantsHelper.ADD_MANAGERIAL_REPORT_PAGE, Constants.ConstantsHelper.ADD_MANAGERIAL_REPORT_WEBPART);
                    createpages(site, Constants.ConstantsHelper.ADD_REPORT_VARIABLE_PAGE, Constants.ConstantsHelper.ADD_REPORT_VARIABLE_WEBPART);
                    createpages(site, Constants.ConstantsHelper.ADD_USER_PAGE, Constants.ConstantsHelper.ADD_USER_WEBPART);
                    createpages(site, Constants.ConstantsHelper.ASTMR_PAGE, Constants.ConstantsHelper.REPRODUCIBILITY_WEBPART);
                    createpages(site, Constants.ConstantsHelper.CHANGE_PASSWORD_PAGE, Constants.ConstantsHelper.CHANGE_PASSWORD_WEBPART);
                    createpages(site, Constants.ConstantsHelper.COMPANY_LIST_PAGE, Constants.ConstantsHelper.COMPANY_LIST_WEBPART);
                    createpages(site, Constants.ConstantsHelper.COPY_FIELD_PAGE, Constants.ConstantsHelper.COPY_FIELD_WEBPART);
                    createpages(site, Constants.ConstantsHelper.COPY_METHOD_PAGE, Constants.ConstantsHelper.COPY_METHOD_WEBPART);
                    createpages(site, Constants.ConstantsHelper.COPY_TEST_PAGE, Constants.ConstantsHelper.COPY_TEST_WEBPART);
                    createpages(site, Constants.ConstantsHelper.DATA_COLLECTION_DATA_ENTRY_PAGE, Constants.ConstantsHelper.DISPLAY_CONTROLS_WEBPART);
                    createpages(site, Constants.ConstantsHelper.DATA_COLLECTION_ENTRY_PAGE, Constants.ConstantsHelper.DATA_COLLECTION_ENTRY_PAGE);
                    createpages(site, Constants.ConstantsHelper.DATA_COLLECTION_RESULT_PAGE, Constants.ConstantsHelper.RESULT_CONTROL_WEBPART);
                    createpages(site, Constants.ConstantsHelper.DATA_COLLECTION_SUBMIT_PAGE, Constants.ConstantsHelper.SUBMIT_CONTROL_WEBPART);
                    createpages(site, Constants.ConstantsHelper.DISPLAY_REPORTS_PAGE, Constants.ConstantsHelper.DISPLAY_REPORTS_WEBPART);
                    createpages(site, Constants.ConstantsHelper.DOWNLOAD_PDF_PAGE, Constants.ConstantsHelper.DOWNLOAD_WEBPART);
                    createpages(site, Constants.ConstantsHelper.EDIT_COMPANY_PAGE, Constants.ConstantsHelper.EDIT_COMPANY_WEBPART);
                    createpages(site, Constants.ConstantsHelper.EDIT_CONTRACT_PAGE, Constants.ConstantsHelper.EDIT_CONTRACT_WEBPART);
                    createpages(site, Constants.ConstantsHelper.EDIT_LAB_PAGE, Constants.ConstantsHelper.EDIT_LAB_WEBPART);
                    createpages(site, Constants.ConstantsHelper.EDIT_MANAGERIAL_REPORT_PAGE, Constants.ConstantsHelper.EDIT_MANAGERIAL_REPORT_WEBPART);
                    createpages(site, Constants.ConstantsHelper.EDIT_PROGRAM_PAGE, Constants.ConstantsHelper.EDIT_PROGRAM_WEBPART);
                    createpages(site, Constants.ConstantsHelper.EDIT_REPORT_VARIABLE_PAGE, Constants.ConstantsHelper.EDIT_REPORT_VARIABLE_WEBPART);
                    createpages(site, Constants.ConstantsHelper.EDIT_USER_PAGE, Constants.ConstantsHelper.EDIT_USER_WEBPART);
                    createpages(site, Constants.ConstantsHelper.LAB_LIST_PAGE, Constants.ConstantsHelper.LAB_LIST_WEBPART);
                    createpages(site, Constants.ConstantsHelper.LATE_REGISTRANTS_PAGE, Constants.ConstantsHelper.LATE_REGISTRANTS_WEBPART);
                    createpages(site, Constants.ConstantsHelper.LATEREGISTRANTS_COMMENT_PAGE, Constants.ConstantsHelper.LATE_REGISTRANT_COMMENT_WEBPART);
                    createpages(site, Constants.ConstantsHelper.MANAGE_FIELD_PAGE, Constants.ConstantsHelper.MANAGE_FIELD_WEBPART);
                    createpages(site, Constants.ConstantsHelper.MANAGE_GROUPCONTRACT_PAGE, Constants.ConstantsHelper.MANAGE_GROUPCONTRACT_WEBPART);
                    createpages(site, Constants.ConstantsHelper.MANAGE_MANAGERIAL_REPORT_PAGE, Constants.ConstantsHelper.MANAGE_MANAGERIAL_WEBPART);
                    createpages(site, Constants.ConstantsHelper.MANAGE_METHOD_PAGE, Constants.ConstantsHelper.MANAGE_METHOD_WEBPART);
                    createpages(site, Constants.ConstantsHelper.MANAGE_PROGRAM_PAGE, Constants.ConstantsHelper.MANAGE_PROGRAM_WEBPART);
                    createpages(site, Constants.ConstantsHelper.MANAGE_RENEWALS_PAGE, Constants.ConstantsHelper.MANAGE_RENEWALS_WEBPART);
                    createpages(site, Constants.ConstantsHelper.MANAGE_TEST_PAGE, Constants.ConstantsHelper.MANAGE_TEST_WEBPART);
                    createpages(site, Constants.ConstantsHelper.MYPROFILE_PAGE, Constants.ConstantsHelper.MYPROFILE_WEBPART);
                    createpages(site, Constants.ConstantsHelper.NOTIFICATION_HISTORY_PAGE, Constants.ConstantsHelper.NOTIFICATION_HISTORY_WEBPART);
                    createpages(site, Constants.ConstantsHelper.PROGRAM_RESULT_UPLOAD_PAGE, Constants.ConstantsHelper.PROGRAM_RESULT_UPLOAD_WEBPART);
                    createpages(site, Constants.ConstantsHelper.PROGRAM_LIST_PAGE, Constants.ConstantsHelper.PROGRAM_LIST_WEBPART);
                    createpages(site, Constants.ConstantsHelper.REMOTE_TEMPLATE_VARIABLE_PAGE, Constants.ConstantsHelper.REMOTE_TEMPLATE_VARIABLE_WEBPART);                    
                    createpages(site, Constants.ConstantsHelper.RUN_REPORT_PAGE, Constants.ConstantsHelper.RUN_REPORT_WEBPART);
                    createpages(site, Constants.ConstantsHelper.SEND_TEST_RENEWALNOTICE_PAGE, Constants.ConstantsHelper.SEND_TEST_RENEWALNOTICE_WEBPART);
                    createpages(site, Constants.ConstantsHelper.UNSUBMITTED_TEST_RESULTS_PAGE, Constants.ConstantsHelper.UNSUBMITTED_RESULTS_WEBPART);
                    createpages(site, Constants.ConstantsHelper.USER_LIST_PAGE, Constants.ConstantsHelper.USER_LIST_WEBPART);
                    createpages(site, Constants.ConstantsHelper.VERSION_DIALOG_PAGE, Constants.ConstantsHelper.VERSION_DIALOG_WEBPART);                    
                    createpages(site, Constants.ConstantsHelper.VIEW_MANAGERIAL_REPORT_PAGE, Constants.ConstantsHelper.VIEW_MANAGERIAL_REPORT_WEBPART);
                    createpages(site, Constants.ConstantsHelper.VIEW_REPORT_PAGE, Constants.ConstantsHelper.VIEW_REPORT_WEBPART);
                    createpages(site, Constants.ConstantsHelper.VIEW_RESULTS_PAGE, Constants.ConstantsHelper.VIEW_RESULT_WEBPART);
                    
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
            try
            {
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
                    if (selectedLayout != null)
                    {
                        PublishingPage newPage = pages.Add(pagename, selectedLayout);
                        newPage.ListItem.Update();
                        newPage.Update();

                        AddWebPartToPage(site, pagename, webpartname);
                        newPage.CheckIn(webpartname + "is added");
                        newPage.ListItem.File.Publish(pagename + " published successfully");
                        Console.WriteLine("Created " + pagename.Replace("_", " "));
                    }
                    else
                    {
                        Console.WriteLine("Page Layout is missing , Page not created ");
                        Console.WriteLine();
                        return;
                    }

                }
                else
                {
                    Console.WriteLine("Already exist " + pagename.Replace("_", " "));
                }
            }
            catch (Exception ex)
            {
            }
            Console.WriteLine();
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
                Console.WriteLine();
            }
            site.AllowUnsafeUpdates = false;
            
        }
        
    }


}

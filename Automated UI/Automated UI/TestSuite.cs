using System;
using System.Security;
using Microsoft.Dynamics365.UIAutomation.Api;
using Microsoft.Dynamics365.UIAutomation.Browser;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automated_UI
{
    /// <summary>
    /// Simple demonstration of the EasyRepro library for Microsoft Dynamics UI Testing.
    /// EasyRepro Version: 9.0.0.1
    /// Reference: https://github.com/Microsoft/EasyRepro/tree/releases/v9.0/Microsoft.Dynamics365.UIAutomation.Sample
    /// </summary>
    [TestClass]
    public class TestSuite
    {
        // Retrieve the url, username&password from three local user Environmental Variables (security). 
        private readonly SecureString username = Environment.GetEnvironmentVariable("CRMUser").ToSecureString();
        private readonly SecureString password = Environment.GetEnvironmentVariable("CRMPassword").ToSecureString();
        private Uri xrmUri = new Uri(Environment.GetEnvironmentVariable("CRMUrl"));

        /// <summary>
        /// Holds some options for the type of browser that will be used for this Test Suite. 
        /// </summary>
        private readonly BrowserOptions options = new BrowserOptions
        {
            BrowserType = BrowserType.Chrome,
            PrivateMode = true,
            FireEvents = true
        };

        /// <summary>
        /// Demonstrate that logging into the CRM instance works.
        /// </summary>
        [TestMethod]
        public void TestLogin()
        {
            // Instantiate a browser session from the BrowserOptions for this test.
            using (var xrmBrowser = new XrmBrowser(options))
            {
                // Login with the retrieved user's account for CRM.
                xrmBrowser.LoginPage.Login(xrmUri, username, password);
                // Display the page for 5 seconds just for demonstration purposes.
                xrmBrowser.ThinkTime(5000);
            }
        }

        /// <summary>
        /// Demonstrate basic CRM navigation and creating a new Organization the way a system user would. 
        /// </summary>
        [TestMethod]
        public void TestCreateOrganization()
        {
            using (var xrmBrowser = new XrmBrowser(options))
            {
                xrmBrowser.LoginPage.Login(xrmUri, username, password);

                // OpenSubArea() is one way to navigate around CRM.
                xrmBrowser.Navigation.OpenSubArea("Grants", "Organizations");
                xrmBrowser.CommandBar.ClickCommand("New");
                // Fields can be assigned in a lot of different ways. This is a simple lookup for the logical field name on the open record and assigning it a value. 
                xrmBrowser.Entity.SetValue("name", "Automated Test Organization");
                // Save the record
                xrmBrowser.CommandBar.ClickCommand("Save & Close");
                //xrmBrowser.Entity.Save(); // Alternative way to save a selected record
                // This will prevent duplicate records from being created. 
                xrmBrowser.Dialogs.DuplicateDetection(true);

                xrmBrowser.ThinkTime(5000);
            }
        }

        /// <summary>
        /// Demonstrate global searching.
        /// </summary>
        [TestMethod]
        public void TestSearch()
        {
            using (var xrmBrowser = new XrmBrowser(options))
            {
                xrmBrowser.LoginPage.Login(xrmUri, username, password);

                xrmBrowser.Navigation.GlobalSearch("Automated Test Organization");

                xrmBrowser.ThinkTime(5000);
            }
        }

        /// <summary>
        /// Demonstrate how to perform a local search, open a record, and delete a record from CRM. 
        /// </summary>
        [TestMethod]
        public void TestDeleteOrganization()
        {
            using (var xrmBrowser = new XrmBrowser(options))
            {
                xrmBrowser.LoginPage.Login(xrmUri, username, password);

                xrmBrowser.Navigation.OpenSubArea("Grants", "Organizations");
                xrmBrowser.Grid.Search("Automated Test Organization");
                xrmBrowser.Grid.OpenRecord(0);
                xrmBrowser.CommandBar.ClickCommand("Delete");
                // This will confirm the delete action (popup window). 
                xrmBrowser.Dialogs.Delete();

                xrmBrowser.ThinkTime(5000);
            }
        }

        /// <summary>
        /// Demonstrate slightly more advanced nagivation of CRM using the API the same way a system user would. 
        /// </summary>
        [TestMethod]
        public void TestNavigation()
        {
            using (var xrmBrowser = new XrmBrowser(options))
            {
                xrmBrowser.LoginPage.Login(xrmUri, username, password);

                xrmBrowser.Navigation.OpenSubArea("Grants", "Funding Opportunities");
                xrmBrowser.Grid.SwitchView("Portal - Apply");
                xrmBrowser.Grid.OpenRecord(0);
                xrmBrowser.Navigation.OpenRelated("Application Section Configurations");

                xrmBrowser.ThinkTime(5000);
            }
        }
    }
}
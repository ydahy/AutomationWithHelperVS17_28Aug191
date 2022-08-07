using NUnit.Framework;
using OpenQA.Selenium;
using AutomationWithHelperVS17_28Aug191.Pageobject;

namespace AutomationWithHelperVS17_28Aug191
{
    [TestFixture]
    public class UnitTest1 : BaseTest
    {
        [Test]

        public void TestMethod1()
        {
            // Arrange
            //var assertBoolean = true;
            HomePage homePage = new HomePage(Driver, Wait);
            Loginpage loginPage = new Loginpage(Driver, Wait);
            LogoutPage logoutPage = new LogoutPage(Driver, Wait);
            
            // Act


            homePage.OpenRigestraion();
            loginPage.LoginStep();

            


            // Assert
            string actualvalue = Driver.FindElements(By.ClassName("ng-scope"))[2].Text;
            Assert.IsTrue(actualvalue.Contains("You're logged in!!"));

            logoutPage.UserLogout();


            Wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(loginPage.Login));
            string Logincontainer = Driver.FindElement(By.ClassName("form-actions")).Text;
            Assert.IsTrue(Logincontainer.Contains("Login"));

            Driver.Dispose();

            // assertBoolean.ShouldSatisfyAllConditions(
            //   () => assertBoolean.ShouldBeTrue(),
            //  () => { if (Settings.FireExceptionOnBrowserConsoleErrors) errors.ShouldBeNullOrEmpty(); });
        }

        /// <summary>
        /// Example of Test Method that takes inputs
        /// Multiple TestCase Attributes for one test method means running the same test method multiple times each with the input provided in the TestCase
        /// </summary>
     /*   [TestCase("text example", Languages.en)]
        [TestCase("text example", Languages.ar)]
        public void TestCase1(string textExample, Languages lang)
        {
            // Arrange
            var assertBoolean = true;
            Settings.Language = lang; // Set the Setting.Language to be used in the LocalizedValueOf function in Common Class when getting from resource files

            // Act


            // Assert
            //GetBrowserErrors(); // Keep this to capture any errors in browser console before assert

            //assertBoolean.ShouldSatisfyAllConditions(
               // () => assertBoolean.ShouldBeTrue(),
               // () => { if (Settings.FireExceptionOnBrowserConsoleErrors) errors.ShouldBeNullOrEmpty(); });
        }*/
    }
}
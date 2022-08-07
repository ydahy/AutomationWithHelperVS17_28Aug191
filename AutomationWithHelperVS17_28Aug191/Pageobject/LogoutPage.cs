using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;

namespace AutomationWithHelperVS17_28Aug191.Pageobject
{
    class LogoutPage : Common
    {

        public LogoutPage(IWebDriver driver, WebDriverWait wait)
        {
            Driver = driver;
            Wait = wait;
        }
        By message = By.XPath("//button[@ng-scope='Auth.login()']");
        By logout = By.LinkText("Logout");
        By Login = By.XPath("//button[@ng-click='Auth.login()']");

        public void UserLogout()

        {

   

            // Assert
            string actualvalue = Driver.FindElements(By.ClassName("ng-scope"))[2].Text;
            Assert.IsTrue(actualvalue.Contains("You're logged in!!"));

            ClickOn(logout, false);
            WaitForPageReadyState();
            Wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(Login));
            string Logincontainer = Driver.FindElement(By.ClassName("form-actions")).Text;
            Assert.IsTrue(Logincontainer.Contains("Login"));

        }
    }
}

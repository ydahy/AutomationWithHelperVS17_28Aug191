using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;

namespace AutomationWithHelperVS17_28Aug191
{
    class Loginpage :Common
    {

        public Loginpage(IWebDriver driver, WebDriverWait wait)
        {
            Driver = driver;
            Wait = wait;
        }

        By Username = By.Id("username");
        By Password = By.Id("password");
        By Username2 = By.Id("formly_1_input_username_0");
        public By Login = By.XPath("//button[@ng-click='Auth.login()']");
        By logout = By.LinkText("Logout");




        public void LoginStep()

        {
            Wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(Username));
            SendKeys(Username, "angular");
            Wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(Password));
            SendKeys(Password, "password");
            Wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(Username2));
            SendKeys(Username2, "Description");
            ClickOn(Login, false);
            WaitForPageReadyState();
            Wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(logout));

        }
    }
}

using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutomationWithHelperVS17_28Aug191.Pageobject
{
    class HomePage : Common

    { 
         public HomePage (IWebDriver driver, WebDriverWait wait)
    {
        Driver = driver;
        Wait = wait;
    }
    
        By Registrationbtn = By.XPath("//h2[contains(text(),'Registration')]");
        public void OpenRigestraion()

        {
            GoToUrl("http://www.way2automation.com/protractor-angularjs-practice-website.html");
            WaitForPageReadyState();
            ScrollTo(Registrationbtn);
            Wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(Registrationbtn));
            ClickOn(Registrationbtn, false);
            Driver.SwitchTo().Window(Driver.WindowHandles[1]);
            WaitForPageReadyState();
        }
    }
}

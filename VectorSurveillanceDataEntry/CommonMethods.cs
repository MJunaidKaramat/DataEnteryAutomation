using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Edge;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;

namespace FP_InterWood
{
    public class GeneralMethodsClass
    {
        public static IWebDriver driver;
        public GeneralMethodsClass()
        {
            //driver = null;
        }
        public void browserSelection(string browserName)
        {
            ChromeOptions options = new ChromeOptions();
            options.AddArguments("--headless");
            if (browserName.Equals("chrome", StringComparison.InvariantCultureIgnoreCase))
            {
                driver = new ChromeDriver();
            }
            else if (browserName.Equals("edge", StringComparison.InvariantCultureIgnoreCase))
            {
                driver = new EdgeDriver();
            }
            driver.Manage().Window.Maximize();

        }
        public void landingPage(string u)
        {
            driver.Url = u;
        }
        public IWebElement findElement(By path)
        {
            IWebElement element = ExplicitWaitElementIsVisible(path);
            return element;
        }

        public void ClickableItem(By path)
        {
            IWebElement element = ExplicitWaitElementIsVisible(path);
            //element.Click();
            Actions action = new Actions(driver);
            action.Click(element).Build().Perform();
        }
 
        public void dropDownItemSelect(By path, string value)
        {
            IWebElement drpDown = findElement(path);
            SelectElement dropDownMenu = new SelectElement(drpDown);
            dropDownMenu.SelectByValue(value);
        }
        
        public void inputText(By path, string data)
        {

            IWebElement txtBox = findElement(path);
            txtBox.SendKeys(data);
        }
        
        public IWebElement ExplicitWaitElementIsVisible(By path)
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(20));
            return wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(path));
        }
        public void scrollPageDown()
        {
            IJavaScriptExecutor scroller = (IJavaScriptExecutor)driver;
            for (int i = 0; i < 1000; i++)
            {
                scroller.ExecuteScript("window.scrollBy(0," + i + ")", "");
            }
        }
        public void scrollToElement(By path)
        {
            IJavaScriptExecutor scroller = (IJavaScriptExecutor)driver;
            IWebElement detectedElement = ExplicitWaitElementIsVisible(path);
            scroller.ExecuteScript("arguments[0].scrollIntoView(true);", detectedElement);
        }
        public void scrollPageUp()
        {
            IJavaScriptExecutor scroller = (IJavaScriptExecutor)driver;
            scroller.ExecuteScript("window.scrollTo(0, document." + "head" + ".scrollHeight);");
        }
        public void hoverAction(By path)
        {
            IWebElement webElement = findElement(path);
            Actions action = new Actions(driver);
            action.MoveToElement(webElement).Build().Perform();
        }
        public void scrollToElementClick(By path)
        {
            IJavaScriptExecutor scroller = (IJavaScriptExecutor)driver;
            IWebElement detectedElement = ExplicitWaitElementIsVisible(path);
            scroller.ExecuteScript("arguments[0].scrollIntoView(true);", detectedElement);
            ClickableItem(path);
        }
        public void closeBrowser()
        {
            driver.Close();
        }
        public void acceptAlert()
        {
            IAlert alert = driver.SwitchTo().Alert();
            alert.Accept();
        }
    }
}

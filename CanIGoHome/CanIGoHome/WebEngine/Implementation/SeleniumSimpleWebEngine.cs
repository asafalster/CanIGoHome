using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.ObjectModel;
using System.Threading;

namespace CanIGoHome
{
    public class SeleniumSimpleWebEngine : IWebEngine
    {
        public bool ConfigureEngine()
        {
            return true;
        }

        public string Search(string user, string pw)
        {
            ChromeDriverService service = ChromeDriverService.CreateDefaultService();
            service.HideCommandPromptWindow = true;
            var options = new ChromeOptions();
            options.AddArgument("headless");

            IWebDriver driver = new ChromeDriver(service, options);
            try
            {
              
                driver.Url = "https://cyberbit.net.hilan.co.il/login";
                //ReadOnlyCollection<string> allWindows = driver.WindowHandles;
                //driver.SwitchTo().Window(allWindows[0]);
                Thread.Sleep(1000);
                IWebElement UserNameInput = driver.FindElement(By.Id("user_nm"));
                UserNameInput.SendKeys(user);

                IWebElement PasswordInput = driver.FindElement(By.Id("password_nm"));
                PasswordInput.SendKeys(pw);

                IWebElement LoginBtn = driver.FindElement(By.TagName("button"));
                LoginBtn.Click();

                Thread.Sleep(2000);
                bool ElementFound = false;
                ReadOnlyCollection<IWebElement> AllTags = driver.FindElements(By.TagName("a"));
                foreach (var webElement in AllTags)
                {
                    if (webElement.GetProperty("innerText").Equals("דיווח ועדכון"))
                    {
                        var exec = (IJavaScriptExecutor)driver;
                        exec.ExecuteScript("arguments[0].click()", webElement);
                        ElementFound = true;
                        break;
                    }
                }

                if (!ElementFound)
                {
                    throw new Exception("Element 'ReportAndUpdate' not found");
                }

                Thread.Sleep(2000);
                IWebElement seletedTD = null;
                try
                {
                    driver.SwitchTo().Frame(0);
                    seletedTD = driver.FindElement(By.CssSelector("td[class='cDIES CSD']"));
                }
                catch
                {
                    //No element is selected
                }

                if (seletedTD == null)
                {

                    //driver.SwitchTo().Frame(0);
                    ElementFound = false;
                    var DateDayStr = DateTime.Now.Day.ToString();
                    ReadOnlyCollection<IWebElement> AllTDs = driver.FindElements(By.ClassName("dTS"));
                    foreach (var webElement in AllTDs)
                    {
                        if (webElement.Text.Equals(DateDayStr))
                        {
                            webElement.Click();
                            ElementFound = true;
                            break;
                        }
                    }

                    if (!ElementFound)
                    {
                        throw new Exception("Element of selected date not found");
                    }


                    IWebElement ShowChosenDaysBtn = driver.FindElement(By.XPath("//input[@value='ימים נבחרים']"));
                    ShowChosenDaysBtn.Click();
                }

                var EnterTimeSpan = driver.FindElement(By.CssSelector("span[class='ROC  gridRowStyle ROC_BG  ItemCell']"));
                if (EnterTimeSpan != null)
                {
                   return EnterTimeSpan.Text;
                }
                else
                {
                    throw new Exception("Element of enter time not found");
                }
            }
            catch (Exception ex)
            {
                return string.Empty;
            }
            finally
            {
                driver.Quit();
            }
        }
    }
}

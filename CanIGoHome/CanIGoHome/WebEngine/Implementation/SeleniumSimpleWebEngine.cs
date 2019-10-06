using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Threading;

namespace CanIGoHome
{
    public class SeleniumSimpleWebEngine : IWebEngine
    {
        private static NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();
        public IDictionary<string,string> Settings { get; set; }
        int _fetchRetries = 1;
        int _defaultDelayBetweenSteps = 1000;
        bool _autoIncreaseDelay = false;
        string _user;
        string _pw;

        public bool ConfigureEngine(IDictionary<string,string> settings)
        {
            logger.Trace($"Enter");
            Settings = settings;

            _fetchRetries = GetSettingsValue<int>("FetchRetries");
            _defaultDelayBetweenSteps = GetSettingsValue<int>("DefaultDelayBetweenSteps");
            _autoIncreaseDelay = GetSettingsValue<bool>("AutoIncreaseDelay");
            _user = GetSettingsValue<string>("User");
            _pw = GetSettingsValue<string>("Password");

            logger.Trace($"Exit");
            return true;
        }

        private T GetSettingsValue<T>(string valueName)
        {
            return (T)Convert.ChangeType(Settings[valueName], typeof(T));
        }

        public string Search()
        {
            logger.Trace($"Enter");
            var delayBetweenSteps = _defaultDelayBetweenSteps;
            ChromeDriverService service = ChromeDriverService.CreateDefaultService();
            service.HideCommandPromptWindow = true;
            var options = new ChromeOptions();
            options.AddArgument("headless");

            IWebDriver driver = new ChromeDriver(service, options);
            for (int i = 0; i < _fetchRetries; i++)
            {
                logger.Debug($"Fetching data - #{i}");
                try
                {
                    logger.Debug($"Navigate to URL");
                    driver.Url = "https://cyberbit.net.hilan.co.il/login";
                    //ReadOnlyCollection<string> allWindows = driver.WindowHandles;
                    //driver.SwitchTo().Window(allWindows[0]);
                    Thread.Sleep(delayBetweenSteps);
                    logger.Debug($"Find user name textbox");
                    IWebElement UserNameInput = driver.FindElement(By.Id("user_nm"));
                    UserNameInput.SendKeys(_user);

                    logger.Debug($"Find pw name textbox");
                    IWebElement PasswordInput = driver.FindElement(By.Id("password_nm"));
                    PasswordInput.SendKeys(_pw);

                    logger.Debug($"Find login button");
                    IWebElement LoginBtn = driver.FindElement(By.TagName("button"));
                    LoginBtn.Click();

                    Thread.Sleep(delayBetweenSteps);
                    logger.Debug($"Find UpdateAndRepport link");
                    bool ElementFound = false;
                    ReadOnlyCollection<IWebElement> AllTags = driver.FindElements(By.TagName("a"));
                    foreach (var webElement in AllTags)
                    {
                        if (webElement.GetProperty("innerText").Equals("דיווח ועדכון"))
                        {
                            logger.Debug($"Click UpdateAndRepport link");
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

                    Thread.Sleep(delayBetweenSteps);
                    logger.Debug($"Find today's date which suppose to be already marked");
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
                        logger.Debug($"Today's date is not marked");

                        //driver.SwitchTo().Frame(0);
                        ElementFound = false;
                        var DateDayStr = DateTime.Now.Day.ToString();
                        logger.Debug($"Find Today's date");
                        ReadOnlyCollection<IWebElement> AllTDs = driver.FindElements(By.ClassName("dTS"));
                        foreach (var webElement in AllTDs)
                        {
                            if (webElement.Text.Equals(DateDayStr))
                            {
                                logger.Debug($"Click on  Today's date");
                                webElement.Click();
                                ElementFound = true;
                                break;
                            }
                        }

                        if (!ElementFound)
                        {
                            throw new Exception("Element of selected date not found");
                        }

                        logger.Debug($"Find Choosen days button");
                        IWebElement ShowChosenDaysBtn = driver.FindElement(By.XPath("//input[@value='ימים נבחרים']"));

                        logger.Debug($"Click Choosen days button");
                        ShowChosenDaysBtn.Click();
                    }

                    logger.Debug($"Find Entrance time");
                    var EnterTimeSpan = driver.FindElement(By.CssSelector("span[class='ROC  gridRowStyle ROC_BG  ItemCell']"));
                    if (EnterTimeSpan != null)
                    {
                        logger.Trace($"Exit");
                        return EnterTimeSpan.Text;
                    }
                    else
                    {
                        throw new Exception("Element of enter time not found");
                    }
                }
                catch (Exception ex)
                {
                    logger.Debug(ex, $"Exception occured");
                    if (_autoIncreaseDelay)
                    {
                        delayBetweenSteps += 500;
                    }
                }
                finally
                {
                    driver.Quit();
                }
            }
            logger.Trace($"Exit");
            return string.Empty;
        }


        [MethodImpl(MethodImplOptions.NoInlining)]
        private string GetCurrentMethod()
        {
            var st = new StackTrace();
            var sf = st.GetFrame(1);

            return sf.GetMethod().Name;
        }
    }
}

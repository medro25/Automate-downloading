using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System.Net.Mail;
using System.Runtime.InteropServices;
using System.Diagnostics;


namespace mDMStoSPOL
{
    class Program
    {

        public static String Filepath;
        public static String objFilepath;
        public static String CopyFailedFile;
        public static String CorruptFile;
        public static String IndexFormDMS;
        public static String Longpath;
        public static DateTime foldercreationTime = new DateTime(636370000000000000);
        public static DateTime FileDownloadTime = new DateTime(636370000000000000);
        public static int NumberOfline;
        public static String OnlyIndexFile = "yes";
        public static String folderURLID;
        public static String versionproblem;
        public static Dictionary<string, string> versionUnconfirmed = new Dictionary<string, string>();
        public static List<string> versioncontrol = new List<string> { };
        static void Main(String[] args)
        {
            ChromeOptions options = new ChromeOptions();
            //String Filepath = @"D:\CPE-HW Management";
            Filepath = args[0];
            options.AddUserProfilePreference("download.default_directory", Filepath);
            options.AddUserProfilePreference("download.prompt_for_download", false);
            options.AddArgument("--safebrowsing-disable-download-protection");
            options.AddUserProfilePreference("download.directory_upgrade", true);
            //options.AddArgument("--disable - web - security");

            if (!Directory.Exists(Filepath))
            {
                Console.WriteLine("That path not exists");
                Directory.CreateDirectory(Filepath);
                Directory.CreateDirectory(Filepath + "\\Log");
                Directory.CreateDirectory(Filepath + "\\Log" + "\\longPath");
            }

            
            // options.AddArgument("no-sandbox");
            ChromeDriver driver = new ChromeDriver(".\\", options, TimeSpan.FromSeconds(1000));
            //SetSafeBrowsing(driver);
            String OBJID = args[1];
            String URL1 = "mdms-ll.int.net.nokia.com/livelink/livelink.exe?func=ll&objId=";
            String URL2 = "&objAction=browse&viewType=1";
            driver.Navigate().GoToUrl("https://" + URL1 + OBJID + URL2);

            Thread.Sleep(3400);
            // Object File initialization and write
            Console.WriteLine("next create file");
            objFilepath = Filepath + "\\log\\" + "MSObjectFile.txt";
            CopyFailedFile = Filepath + "\\log\\" + "CopyFailedFile.txt";
            CorruptFile = Filepath + "\\log\\" + "CorruptFile.txt";
            IndexFormDMS = Filepath + "\\log\\" + "IndexFormDMS.csv";
            Longpath = Filepath + "\\log\\" + "Longpath.txt";
            versionproblem= Filepath + "\\log\\" + "versionproblem.txt";
            File.AppendAllText(versionproblem, "All the problematic files in version control will be written here, please check manyally for copyfailed/unconfirme files" + Environment.NewLine);
            File.AppendAllText(objFilepath, "Here object ID will be listed" + DateTime.UtcNow + Environment.NewLine);
            Console.WriteLine("test2");
            File.AppendAllText(CopyFailedFile, "Here Failed file ID will be listed" + DateTime.UtcNow + Environment.NewLine);
            File.AppendAllText(CorruptFile, "Here object ID will be listed" + DateTime.UtcNow + Environment.NewLine);
            File.AppendAllText(IndexFormDMS, "Element-OBJID,Type,Element Name,Element Absolute Path , Size,URL," + DateTime.UtcNow + "," + Environment.NewLine);
            File.AppendAllText(Longpath, "Here object ID will be listed to be deleted " + DateTime.UtcNow + Environment.NewLine);
            
            driver.Manage().Window.Maximize();

            //driver.FindElement(By.XPath("//*[contains(@id,'advancedButton')]")).Click();
            //driver.FindElement(By.XPath("//*[contains(@id,'exceptionDialogButton')]")).Click();
            DownloadFileLink(driver, Filepath, false, Filepath, folderURLID);    // In the downloadfileLink method false is Multpage variable
            
            Console.WriteLine("Download completed , Please check before closing browser.");
            Console.WriteLine("Trying to confirm all unconfirmed downloads");
            ConfirmDownloadFile(driver,Filepath);
            Version_unconfiremd(Filepath);
            Console.WriteLine("Done");
            //  driver.Close();
        }
  
        public static IList<IWebElement> shadow(IWebElement s,IWebDriver d)
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)d;
            return (IList<IWebElement>)js.ExecuteScript("return arguments[0].shadowRoot.children", s);
        }

        //sets browser to safemode to have less provlematic files
        public static void SetSafeBrowsing(IWebDriver dr)
        { 
            dr.Navigate().GoToUrl("chrome://settings/security");
            Thread.Sleep(400);
            IWebElement s = dr.FindElement(By.TagName("settings-ui"));
            IList<IWebElement> slist = shadow(s,dr);
            IWebElement sm=slist[5].FindElement(By.TagName("settings-main"));
            IList<IWebElement> smlist = shadow(sm, dr);
            IList<IWebElement> sbplist = shadow(smlist[3], dr);
            IWebElement spp = sbplist[1].FindElement(By.TagName("settings-privacy-page"));
            IList<IWebElement> sppList = shadow(spp, dr);
            IWebElement ssp = sppList[3].FindElement(By.TagName("settings-subpage")).FindElement(By.TagName("settings-security-page"));
            IList<IWebElement> sspList = shadow(ssp, dr);
            IWebElement safeBr = sspList[2].FindElement(By.Id("safeBrowsingStandard"));
            dr.Manage().Window.Maximize();
            Thread.Sleep(400);
            ((IJavaScriptExecutor)dr).ExecuteScript("arguments[0].scrollIntoView(true);", safeBr);
            safeBr.Click();
            Thread.Sleep(100);
        }
 

        //confirms all the unconfirmed files in chrome downloads, uses keyboard to accept alert given by google chrome
        [DllImport("user32.dll")]
        public static extern bool SendMessage(IntPtr hWnd, uint Msg, int wParam, int lParam);
        [DllImport("user32.dll")]
        public static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);
        public static void ConfirmDownloadFile(IWebDriver driver ,String filepath)
            {
            driver.Navigate().GoToUrl("chrome://downloads/");
            string[] fileEntries = Directory.GetFiles(filepath);
            Console.WriteLine(fileEntries.Length);
            int unconfirmed = 0;
            int confirmed = 0;
            foreach (String fileEntry in fileEntries)
            {
                if (fileEntry.Contains(".crdownload"))
                {
                    unconfirmed++;
                }
            }
            Console.WriteLine(unconfirmed);
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            if (unconfirmed > 0) {
                Console.WriteLine("Unconfirmed files found, confirm them?y/n");
                string con = Console.ReadLine();
                if(con=="y" || con == "yes")
                {
                    while (confirmed < unconfirmed)
                    {

                        Console.WriteLine("confirming");
                        IWebElement find = driver.FindElement(By.CssSelector("downloads-manager"));
                        IList<IWebElement> list = (IList<IWebElement>)js.ExecuteScript("return arguments[0].shadowRoot.children", find);
                        IList<IWebElement> downloads = list[3].FindElements(By.Id("downloadsList"));
                        IList<IWebElement> buttons = downloads[0].FindElements(By.TagName("downloads-item"));
                        Actions actions = new Actions(driver);
                        Console.WriteLine(buttons.Count + " downloads found");
                        IList<Process> processes = Process.GetProcessesByName("chrome");

                        for (int k = 0; k < buttons.Count; k++)
                        {
                            try
                            {
                                IList<IWebElement> buttonblock = (IList<IWebElement>)js.ExecuteScript("return arguments[0].shadowRoot.children", buttons[k]);
                                Console.WriteLine("id: " + buttonblock[2].GetAttribute("id"));
                               
                                actions.MoveToElement(buttonblock[2]).Perform();
                                if (buttonblock[2].FindElement(By.TagName("iron-icon")).GetAttribute("outerHTML").ToString().Contains("red"))
                                {
                                    IList<IWebElement> b = buttonblock[2].FindElements(By.TagName("cr-button"));
                                    IList<IWebElement> a = buttonblock[2].FindElements(By.TagName("a"));
                                    foreach (IWebElement button in b)
                                    {
                                        Console.WriteLine("Running");
                                        Console.WriteLine(button.Displayed);
                                        if (button.GetAttribute("focus-type") == "save"&& button.Displayed)
                                        {
                                            foreach (IWebElement aa in a)
                                            {
                                                if (aa.GetAttribute("id") == "show")
                                                {
                                                    Console.WriteLine(aa.GetAttribute("title"));
                                                }
                                            }
                                            Console.WriteLine("button found");
                                            const int WM_SYSKEYDOWN = 0x0104;
                                            const int VK_KEY_RIGHT = 0x27;
                                            const int VK_KEY_ENTER = 0x0D;

                                            Thread.Sleep(10);
                                            if (button.Enabled)
                                            {
                                                Console.WriteLine("clicking the button");
                                                actions.MoveToElement(button).Perform();
                                                js.ExecuteScript("arguments[0].scrollIntoView(true);", button);
                                                button.Click();
                                            }

                                            Console.WriteLine("clicked the button");

                                            IntPtr windowHandle = processes[0].MainWindowHandle;
                                            SendMessage(windowHandle, WM_SYSKEYDOWN, VK_KEY_RIGHT, 0);
                                            SendMessage(windowHandle, WM_SYSKEYDOWN, VK_KEY_ENTER, 0);
                                            confirmed++;
                                            Console.WriteLine("out of foreachloop");

                                        }

                                        if (button.GetAttribute("id") == "remove" && button.Displayed)
                                        {
                                            Console.WriteLine("removing");
                                            js.ExecuteScript("arguments[0].scrollIntoView(true);", button); ;
                                            actions.MoveToElement(button).Perform();
                                            Thread.Sleep(10);
                                            if (button.Enabled)
                                            {
                                                button.Click();
                                            }
                                        }
                                        actions.MoveByOffset(0, 300);
                                        actions.Perform();
                                    }
                                }
                                else
                                {
                                    Console.WriteLine("downloaded");
                                    try
                                    {
                                        IWebElement button = buttonblock[2].FindElement(By.TagName("cr-icon-button"));

                                        if (button.Enabled && button.Displayed)
                                        {
                                            js.ExecuteScript("arguments[0].scrollIntoView(true);", button);
                                            button.Click();
                                        }
                                        Console.WriteLine("removed");
                                    }
                                    catch (Exception e)
                                    {
                                        Console.WriteLine(e);
                                    }
                                }

                            }
                            catch (Exception e)
                            {

                                Console.WriteLine("Exception: " + e);
                                driver.Navigate().Refresh();
                                find = driver.FindElement(By.CssSelector("downloads-manager"));
                                list = (IList<IWebElement>)js.ExecuteScript("return arguments[0].shadowRoot.children", find);
                                downloads = list[3].FindElements(By.Id("downloadsList"));
                                buttons = downloads[0].FindElements(By.TagName("downloads-item"));
                                actions = new Actions(driver);
                                Console.WriteLine(buttons.Count + " downloads found");
                            }
                        }
                    }
                }
            
            }
        }




        // downloads one file, given list of links and current index of list, code moved out of the for loop in DownloadFileLink
        public static void DownloadFile(IWebDriver driver, IList<IWebElement> DownloadLink, int i, string sourcepath,string path, WebDriverWait wait)
        {
            String currenturl = driver.Url;
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(120);
            DownloadLink = driver.FindElements(By.XPath("//a[text()='Download']"));
            Console.WriteLine("Download block: " + DownloadLink.Count);
            Thread.Sleep(600);

            DownloadLink = driver.FindElements(By.XPath("//a[text()='Download']"));
            // wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.VisibilityOfAllElementsLocatedBy(By.XPath("//a[text()='Download']")));
            Console.WriteLine("This is value i " + DownloadLink[i]);

            wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(DownloadLink[i]));
            // Thread.Sleep(200);
            Console.WriteLine("wait over  " + DownloadLink.Count);
            DownloadLink = driver.FindElements(By.XPath("//a[text()='Download']"));
            String Downloadfile = DownloadLink[i].GetAttribute("href").ToString();
            var SrcFileName = path + DownloadLink[i].Text;
            // download file click on download button
            if (Downloadfile.Contains("download"))
            {
                if (readObjIDFile(Downloadfile, objFilepath))
                {
                    Console.WriteLine(Downloadfile + "\t  File has been skipped");
                    Console.WriteLine("InSide if block :" + i);

                }
                else
                {
                    Console.WriteLine("InSide else block :" + i);
                    Thread.Sleep(400);
                    DownloadLink = driver.FindElements(By.XPath("//a[text()='Download']"));
                    Console.WriteLine("element ID before wait :" + DownloadLink[i]);
                    wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(DownloadLink[i]));
                    Thread.Sleep(300);
                    DownloadLink = driver.FindElements(By.XPath("//a[text()='Download']"));
                    Thread.Sleep(300);
                    // driver.Manage().Timeouts().AsynchronousJavaScript = TimeSpan.FromSeconds(2000);
                    // wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(DownloadLink[i]));
                    // wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.VisibilityOfAllElementsLocatedBy(By.XPath("//a[text()='Download']")));
                    driver.Manage().Timeouts().PageLoad = TimeSpan.FromMinutes(80);
                    DownloadLink[i].Click();
                    Thread.Sleep(400);  //8Aug 400 to 200

                    Console.WriteLine("file downloaded" + DownloadLink[i]);
                    //Validate and report corrupt file
                    if (driver.FindElements(By.XPath("//*[contains(@id,'DivLocationSelectTitle')]")).Count == 0)
                    {
                        Console.WriteLine("Corrupt file found " + "\t" + Downloadfile);
                        driver.Navigate().Back();
                        driver.Manage().Timeouts().AsynchronousJavaScript = TimeSpan.FromSeconds(200);
                        Console.WriteLine("Corrupt file has been skiped " + DownloadLink[i] + "\t" + Downloadfile);
                        WriteObjIDtoFile(Downloadfile);

                        if (DownloadLink.Count > 25)
                        {
                            driver.FindElement(By.XPath("//*[contains(@id,'selectedBrowsePageSize')]")).Click();
                            Thread.Sleep(400);
                            driver.FindElement(By.XPath("//*[contains(@id,'BrowsePageSize125DivId')]")).Click();
                            Thread.Sleep(20000);
                        }


                        DownloadLink = driver.FindElements(By.XPath("//a[text()='Download']"));
                        //private static void createIndex(string linkobj, string typeOfFile, String filepath, String Filename, String ElementOBJID, String ObjectSize)
                        createIndex("Corrupt File", DownloadLink[i].Text.ToString(), path, "0", Downloadfile);
                        File.AppendAllText(CorruptFile, Downloadfile + "; " + path + "\tCourrpt file" + Environment.NewLine);
                        //write objID+corrupt skip copy commb
                        DownloadLink = driver.FindElements(By.XPath("//a[text()='Download']"));
                        Console.WriteLine(" number of files " + DownloadLink);
                        DownloadLink = driver.FindElements(By.XPath("//a[text()='Download']"));
                        driver.Manage().Timeouts().AsynchronousJavaScript = TimeSpan.FromSeconds(200);
                    }
                    //Validate and report corrupt file
                    // if(CheckSourceFile)

                }


                

            }
        }

        public static void DownloadFileLink(IWebDriver driver, string path, Boolean multipage, String SourcePath, String objFolderID)
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromMinutes(5));

            //Download Files
            try
            {
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
                //Increase page size from drop down from 25 to 125 at bottom of page
                int Pagesize = driver.FindElements(By.XPath("//*[contains(@id,'selectedBrowsePageSize')]")).Count;
                int pagecount = driver.FindElements(By.ClassName("browseItemName")).Count;
                Console.WriteLine("Pagesize: " + Pagesize);

                Console.WriteLine("pagecount: " + pagecount);

                if (Pagesize != 0 && multipage == false)
                {
                    driver.FindElement(By.XPath("//*[contains(@id,'selectedBrowsePageSize')]")).Click();
                    Thread.Sleep(400);
                    driver.FindElement(By.XPath("//*[contains(@id,'BrowsePageSize125DivId')]")).Click();
                    Thread.Sleep(400);

                    do
                    {
                        pagecount = driver.FindElements(By.ClassName("browseItemName")).Count;
                        Console.WriteLine("pagecount in side dowhile loop : " + pagecount);

                    } while (pagecount < 25);
                    //driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(45);- 08-Aug
                    Thread.Sleep(2000);
                }
                //Download click and download file in current folder
                //debug
                else
                {
                    Thread.Sleep(2000);
                    var count = driver.FindElements(By.XPath("//a[text()='Download']")).Count();
                    Console.WriteLine("Download item count: " + count);
                }


                //SkipUnsecureFile(driver,path);

                if (driver.FindElements(By.XPath("//a[text()='Download']")).Count != 0)
                {

                    IList<IWebElement> DownloadLink = driver.FindElements(By.XPath("//a[text()='Download']"));

                    Thread.Sleep(600);
                    for (int i = 0; i < DownloadLink.Count; i++)
                    {
                        if (DownloadLink[i].GetAttribute("href").ToString().Contains("Action=download"))
                        {
                            DownloadFile(driver, DownloadLink,i, SourcePath, path, wait);
                        }
                        else
                        {
                            Console.WriteLine("Skipped " + DownloadLink[i]);
                        }
                            //String userConfirm = "No";
                            //checkdownload(driver, userConfirm);
                    }
                }

                // Check here keep

            }
            catch (NoSuchElementException e)
            {
                Console.WriteLine(e);
            }
            //Download folders inside folders 
            try
            {

                IList<IWebElement> AllLink = driver.FindElements(By.ClassName("browseItemName"));
                /* IList<IWebElement> AllLink1 = driver.FindElements(By.ClassName("browseItemName"));
                 IList<IWebElement> AllLink2 = driver.FindElements(By.ClassName("catalog"));
                 IList<IWebElement> AllLink = AllLink1.Union(AllLink2).ToList();*/
                IList<IWebElement> AllLinkItemSize = driver.FindElements(By.ClassName("browseItemSize")); ///Get Item Size.

                //Console.WriteLine("checking alllink");

                List<string> hyperlink = new List<string>(),
                items = new List<string>();
                List<string> ItemSize = new List<String>();


                // Cretae Item List for all valid items and remove corrupt folders 
                for (int i = 0; i < AllLink.Count; i++)
                {
                    //validate folder downloaded or need  to skip

                    String tempHyperlink = AllLink[i].FindElement(By.TagName("a")).GetAttribute("href");
                    Console.WriteLine("URL :" + tempHyperlink);
                    if (folderURLID == getObjID(tempHyperlink).ToString())
                    {
                        Console.WriteLine("Not a real folder - Loop found");
                        File.AppendAllText(CorruptFile, "Edit Minor :" + tempHyperlink + "; " + path + "\""+AllLink[i].Text.Trim()+"\t Edit Minor "+ Environment.NewLine);
                    }


                    if (readObjIDFile(tempHyperlink, objFilepath))
                    {
                        Console.WriteLine(AllLink[i].Text.Trim() + " folder has been  skiped");
                    }

                    /* else if (path == Filepath && !tempHyperlink.Contains("Action=browse"))

                     {
                         Console.WriteLine(AllLink[i].Text.Trim() + " Source Folder is same as targetfolder :Skipped file");
                         WriteObjIDtoFile(tempHyperlink);
                         // private static void createIndex(string linkobj, string typeOfFile, String filepath, String Filename, String ElementOBJID, String ObjectSize)
                         createIndex("Parent-File", AllLink[i].Text.Trim(), path,AllLinkItemSize[i].Text.ToString(), tempHyperlink);
                     }*/
                    else
                    {

                        if (tempHyperlink.StartsWith("https://mdms-ll.int.net.nokia.com/") && !tempHyperlink.Contains("objaction=open"))
                        {
                            hyperlink.Add(AllLink[i].FindElement(By.TagName("a")).GetAttribute("href"));
                            items.Add(AllLink[i].Text.Trim());
                            ItemSize.Add(AllLinkItemSize[i].Text.ToString());
                        }
                        else if (tempHyperlink.Contains("objaction=open"))
                        {
                            Console.WriteLine("Short cut for other folder found   " + tempHyperlink + "\t" + AllLink[i].Text.Trim());
                            File.AppendAllText(CorruptFile, tempHyperlink + "; " + path + "\"" + AllLink[i].Text.Trim() + "\tShortCut for other folder " + Environment.NewLine);
                        }

                        else
                        {
                            Console.WriteLine("Link  has been skiped as link is not accessible   " + tempHyperlink + "\t" + AllLink[i].Text.Trim());
                            File.AppendAllText(CorruptFile, tempHyperlink + "; " + path + "\"" + AllLink[i].Text.Trim() + "\tLink Not accessible file " + Environment.NewLine);

                        }


                    }

                }

                //adding same url in item as it have multiple pages to check
                if (driver.Url.ToString().Contains("1__125_"))
                {
                    File.AppendAllText(CorruptFile, "File size is euqal to  125 ; " + driver.Url.ToString() + ";" + path + Environment.NewLine);
                    Console.WriteLine("last folder where beyond 125 items");
                    hyperlink.Add(driver.Url.ToString());

                    items.Add("last folder where beyond 125 items");
                    //var nextpage = driver.FindElement(By.XPath("//*[@id='PageNextImg']"));
                    //if (nextpage.Enabled) { nextpage.Click(); DownloadFileLink(driver, path); }

                }
                // Copy downloded file from source to destination folder
                for (int i = 0; i < items.Count; i++)
                {
                    //copy file from source to destionation

                    Boolean iffolder = (hyperlink[i].Contains("Action=browse") | hyperlink[i].Contains("Action=open&viewType")) && !hyperlink[i].Contains(".html") && !hyperlink[i].Contains(".htm");
                    String FileObjID = getObjID(hyperlink[i]);
                    if (!iffolder)
                    {
                        Console.WriteLine("move file from source to destination : " + items[i]);

                        String fileName = filenameCorerction(items[i]);
                        // foldercreationTime
                        string sourcePath = @Filepath;
                        int counter = 0;
                        // Use Path class to manipulate file and directory paths.
                        string sourceFile = System.IO.Path.Combine(sourcePath, fileName);
                        // FileDownloadTime = File.GetCreationTime(sourceFile);
                        string destFile = System.IO.Path.Combine(path, fileName);
                        string OriginalsourceFile = System.IO.Path.Combine(sourcePath, fileName);
                        string OriginaldestFile = System.IO.Path.Combine(path, fileName);

                        //sourceFile = removeExtention(sourceFile); //Why??
                        //destFile = removeExtention(destFile); // Why
                        // Check file available at source 
                        String Unsecure_extention = Path.GetExtension(fileName);
                        String RenameFile = "NOK";
                        do
                        {


                            if (File.Exists(OriginalsourceFile + "*.crdownload"))
                            {
                                Thread.Sleep(200);
                            }


                            else if (File.Exists(OriginalsourceFile))
                            {
                                sourceFile = OriginalsourceFile;
                                destFile = OriginaldestFile;
                                break;
                            }

                            else if (Unsecure_extention == ".swf")
                            {

                                Console.WriteLine(" SWF file has been skipped , moveing to next file ." + destFile);
                                sourceFile = System.IO.Path.Combine(sourcePath, fileName);
                                destFile = System.IO.Path.Combine(path, fileName);
                                counter = 200;
                                break;

                            }
                            else if (File.Exists(sourceFile + ".ppta"))
                            {
                                sourceFile += ".ppta";
                                destFile += ".ppta";
                                break;
                            }

                            else if (File.Exists(sourceFile + ".DWG"))
                            {
                                sourceFile += ".DWG";
                                destFile += ".DWG";
                                break;
                            }
                            else if (File.Exists(sourceFile + ".dwg"))
                            {
                                sourceFile += ".dwg";
                                destFile += ".dwg";
                                break;
                            }

                            else if (File.Exists(sourceFile + ".tif"))
                            {
                                sourceFile += ".tif";
                                destFile += ".tif";
                                break;
                            }
                            else if (File.Exists(sourceFile + ".html"))
                            {
                                sourceFile += ".html";
                                destFile += ".html";
                                break;
                            }
                            else if (File.Exists(sourceFile + "_"))
                            {
                                sourceFile += "_";
                                destFile += "_";
                                break;
                            }
                            else if (File.Exists(sourceFile + ".zip"))
                            {
                                sourceFile += ".zip";
                                destFile += ".zip";
                                break;
                            }

                            else if (File.Exists(sourceFile + ".swf"))
                            {
                                sourceFile += ".swf";
                                destFile += ".swf";
                                break;
                            }
                            else if (File.Exists(sourceFile + ".fpga"))
                            {
                                sourceFile += ".fpga";
                                destFile += ".fpga";
                                break;
                            }

                            else if (File.Exists(sourceFile))
                            {
                                break;
                            }
                            else if (File.Exists(sourceFile + ".doc"))
                            {
                                sourceFile += ".doc";
                                destFile += ".doc";
                                break;
                            }
                            else if (File.Exists(sourceFile + ".rar"))
                            {
                                sourceFile += ".rar";
                                destFile += ".rar";
                                break;
                            }
                            else if (File.Exists(sourceFile + ".doc" + ".zip"))
                            {
                                sourceFile += ".doc" + ".zip";
                                destFile += ".doc" + ".zip";
                                break;
                            }
                            else if (File.Exists(sourceFile + ".docx"))
                            {
                                sourceFile += ".docx";
                                destFile += ".docx";
                                break;
                            }
                            else if (File.Exists(sourceFile + ".ppt"))
                            {
                                sourceFile += ".ppt";
                                destFile += ".ppt";
                                break;
                            }
                            else if (File.Exists(sourceFile + ".pdf"))
                            {
                                sourceFile += ".pdf";
                                destFile += ".pdf";
                                break;
                            }


                            else if (File.Exists(sourceFile + ".xls"))
                            {
                                sourceFile += ".xls";
                                destFile += ".xls";
                                break;
                            }
                            else if (File.Exists(sourceFile + ".tgz"))
                            {
                                sourceFile += ".tgz";
                                destFile += ".tgz";
                                break;
                            }
                            else if (File.Exists(sourceFile + ".xlsx"))
                            {
                                sourceFile += ".xlsx";
                                destFile += ".xlsx";
                                break;
                            }

                            else if (File.Exists(sourceFile + ".txt"))
                            {
                                sourceFile += ".txt";
                                destFile += ".txt";
                                break;
                            }
                            else if (File.Exists(sourceFile + ".csv"))
                            {
                                sourceFile += ".csv";
                                destFile += ".csv";
                                break;
                            }
                            else if (File.Exists(sourceFile + ".xlsm"))
                            {
                                sourceFile += ".xlsm";
                                destFile += ".xlsm";
                                break;
                            }
                            else if (File.Exists(sourceFile + ".rtf"))
                            {
                                sourceFile += ".rtf";
                                destFile += ".rtf";
                                break;
                            }
                            else if (File.Exists(sourceFile + ".htm"))
                            {
                                sourceFile += ".htm";
                                destFile += ".htm";
                                break;
                            }

                            else if (File.Exists(sourceFile + ".pptx"))
                            {
                                sourceFile += ".pptx";
                                destFile += ".pptx";
                                break;
                            }
                            else if (File.Exists(sourceFile + ".gz"))
                            {
                                sourceFile += ".gz";
                                destFile += ".gz";
                                break;
                            }
                            else if (File.Exists(sourceFile + ".mpp"))
                            {
                                sourceFile += ".mpp";
                                destFile += ".mpp";
                                break;
                            }
                            else if (File.Exists("__" + sourceFile))
                            {
                                sourceFile = "__" + sourceFile;
                                destFile = "__" + sourceFile;
                                break;
                            }
                            else if (File.Exists("_" + sourceFile))
                            {
                                sourceFile = "_" + sourceFile;
                                destFile = "_" + sourceFile;
                                break;
                            }


                            else if (sourceFile == destFile)
                            {
                                Console.WriteLine("Sorce File: {0}  and Destination File : {1} is same", sourceFile, destFile);
                                break;
                            }

                            else if (counter == 100 & RenameFile == "NOK")
                            {
                                Console.WriteLine("Wait{0} fileName :  {1} ", counter, sourceFile);
                                string[] fileEntries = Directory.GetFiles(sourcePath);


                                foreach (string fileNameS in fileEntries)
                                {
                                    //String updatedFileNAme = Regex.Replace(fileNameS, @"\s+", " ");
                                    try
                                    {

                                        if (!fileNameS.Contains(".crdownload"))
                                        {


                                            System.IO.File.Move(fileNameS, Regex.Replace(fileNameS, @"\s+", " "), true);
                                            Console.WriteLine("Correction of file name" + fileNameS);
                                        }
                                    }
                                    catch (Exception e)
                                    {
                                        Console.WriteLine(e);
                                    }

                                }
                                counter += 1;
                                RenameFile = "OK";
                            }

                            else if (File.Exists(sourceFile + ".msg"))
                            {
                                sourceFile += ".msg";
                                destFile += ".msg";

                                break;

                            }
                            else if (File.Exists(sourceFile + ".til"))
                            {
                                sourceFile += ".til";
                                destFile += ".til";

                                break;

                            }
                            else if (File.Exists(sourceFile + ".xml"))
                            {
                                sourceFile += ".xml";
                                destFile += ".xml";

                                break;

                            }
                            else if (File.Exists(sourceFile + ".bin"))
                            {
                                sourceFile += ".bin";
                                destFile += ".bin";
                                break;
                            }
                            else if (Unsecure_extention == ".js")
                            {
                                Console.WriteLine(" JS file has been skipped , moveing to next file ." + destFile);
                                sourceFile = System.IO.Path.Combine(sourcePath, fileName);
                                destFile = System.IO.Path.Combine(path, fileName);
                                counter = 200;
                                break;

                            }
                            else if (Unsecure_extention == ".msg")
                            {
                                Console.WriteLine(" msg file has been skipped , moveing to next file ." + destFile);
                                sourceFile = System.IO.Path.Combine(sourcePath, fileName);
                                destFile = System.IO.Path.Combine(path, fileName);
                                counter = 200;
                                break;

                            }
                            else if (Unsecure_extention == ".xml")
                            {
                                Console.WriteLine(" xml file has been skipped , moveing to next file ." + destFile);
                                sourceFile = System.IO.Path.Combine(sourcePath, fileName);
                                destFile = System.IO.Path.Combine(path, fileName);
                                counter = 200;
                                break;

                            }
                            else if (counter == 200)
                            {
                                Console.WriteLine("Wait{0} fileName :  {1} ", counter, sourceFile);
                                Console.WriteLine("file  has been skipped , moveing to next file ." + destFile);
                                Console.WriteLine("SourceFile Name :" + sourceFile + " And OriginalFilename :" + OriginalsourceFile);
                                sourceFile = System.IO.Path.Combine(sourcePath, fileName);
                                destFile = System.IO.Path.Combine(path, fileName);

                                break;
                            }

                            else
                            {
                                Thread.Sleep(200);
                                counter += 1;

                            }

                        } while (true);

                        if (File.Exists(sourceFile))
                        {
                            FileDownloadTime = File.GetCreationTime(sourceFile);
                            if (foldercreationTime > FileDownloadTime)
                            {
                               Console.WriteLine(" Duplicate file available skip and add to copyfailedbat : " + sourceFile);
                                //sourceFile = CheckSourceFile(sourceFile, destFile, foldercreationTime);

                            }
                        }

                        if (destFile.Length > 260)
                        {
                            // Thread.Sleep(1);
                            string timestamp = DateTime.UtcNow.ToString("yyyyMMddHHmmss");
                            string NewFileName = timestamp + "-" + fileName;
                            string longpathfile = sourcePath + "\\Log\\" + "LongPath\\" + NewFileName;

                            //if file exits at destination alreadythen rename before moving

                            if (File.Exists(sourceFile) & foldercreationTime < FileDownloadTime & !File.Exists(longpathfile))
                            {

                                System.IO.File.Move(sourceFile, longpathfile, false);
                                //File.AppendAllText(CorruptFile, "OriginalFile :" + fileName + ",NewFIleNAme :" + NewFileName + " ,Original Path is : " + destFile + Environment.NewLine);
                                File.AppendAllText(Longpath, "move \"" + longpathfile + "*\" \"" + destFile + "\"" + Environment.NewLine);
                            }
                            //else if file move source file onlyif available 
                            else
                            {
                                File.AppendAllText(CopyFailedFile, "move \"" + sourceFile + "\" \"" + longpathfile + "\"" + Environment.NewLine);
                                File.AppendAllText(Longpath, "move \"" + longpathfile + "*\" \"" + destFile + "\"" + Environment.NewLine);
                                //File.AppendAllText(CorruptFile, "OriginalFile :" + fileName + ",NewFIleNAme :" + NewFileName + " ,Original Path is : " + destFile + Environment.NewLine);

                            }
                            Console.WriteLine("file with longPath  found  : " + fileName + "\t" + destFile);

                        }
                        else
                        {
                            FileDownloadTime = File.GetCreationTime(sourceFile);
                            //  if (sourceFile != destFile & counter != 200 & File.Exists(destFile) == false & foldercreationTime < FileDownloadTime & !destFile.EndsWith(".html")) //Correction on 09/02  to handle html files .
                            if (sourceFile != destFile & counter != 200 & File.Exists(destFile) == false & foldercreationTime < FileDownloadTime)
                            {
                                Console.WriteLine(destFile + "\tcopying......");
                                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(130);
                                Console.WriteLine(" This is a " + sourceFile);
                                System.IO.File.Copy(sourceFile, destFile, true);
                                File.Delete(sourceFile);
                            }
                            else if (Path.GetExtension(destFile) == ".swf")
                            {
                                Console.WriteLine("This file is .swf file");

                            }
                            //Create move command as file has not moved to destination location
                            else
                            {

                                File.AppendAllText(CopyFailedFile, "move \"" + sourceFile + "\" \"" + path + '"' + '\t' + "%=" + FileObjID + "=%" + Environment.NewLine);

                            }
                        }
                        //Write object ID to Objid File
                        //if (!fileName.EndsWith(".html") | !fileName.EndsWith(".htm"))
                        if (folderURLID == FileObjID)
                        {
                            Console.WriteLine(" File ObjID is same as folder OBJiD so not writing to MSOBJ file ....." + FileObjID);
                        }
                        else
                        {
                            WriteObjIDtoFile(hyperlink[i]);
                            // private static void createIndex(string linkobj, string typeOfFile, String filepath, String Filename, String ElementOBJID, String ObjectSize)
                            createIndex("File", fileName, path, AllLinkItemSize[i].Text.ToString(), hyperlink[i]);
                        }


                    }
                    


                    else
                    {
                        Console.WriteLine("Folder need to process later : " + items[i]);
                    }

                }

                for(int i=0; i < hyperlink.Count; i++)
                {
                    Download_versions(hyperlink[i], driver, SourcePath, path);
                   Thread.Sleep(100);
                }

                //Go to child folder , create folder and navigate to folder
                //:CON, PRN, AUX, NUL, COM1, COM2, COM3, COM4, COM5, COM6, COM7, COM8, COM9, LPT1, LPT2, LPT3, LPT4, LPT5, LPT6, LPT7, LPT8, and LPT9

                for (int i = 0; i < items.Count; i++)
                {

                    if ((hyperlink[i].Contains("Action=browse") | hyperlink[i].Contains("Action=open&viewType=1")) && !hyperlink[i].ToString().Contains("1__125_")
                            && !items[i].EndsWith(".html") && !items[i].EndsWith(".htm") && !hyperlink[i].Contains("fetch/2000"))
                    //neha 
                    {
                        Console.WriteLine("Inside a folder : " + items[i]);
                        //create folder
                        //Console.WriteLine(FolderName + "\tFolder getting downloaded");
                        string FolderName = System.IO.Path.Combine(path, items[i].Replace("/", " ").
                        Replace("?", "_").Replace(">", "").Replace("<", "").Replace("\"", "_").Replace("*", "").Replace("|", "_").
                        Replace("10_Source_In_Development ", "Dev").Replace("...", "").Replace("..", "").

                        Replace("20_Source_In_Validation", "Validated").
                        Replace("10_Source_In_Review ", "In_Review").
                        Replace("Instructor material", "InstructorData").
                        Replace("Instructor Community", "Inst Comm").
                        Replace("Trainee material", "Trainee").
                        Replace("10_Analysis", "Analysis").
                        Replace("30_Implementation", "Implement").Trim()

                        );

                        //Check foldername if a reserve work                             



                        Directory.CreateDirectory(FolderName);
                        foldercreationTime = Directory.GetCreationTime(FolderName);
                        driver.Navigate().GoToUrl(hyperlink[i]);
                        folderURLID = getObjID(hyperlink[i]).ToString();
                        wait.Until(d => ((IJavaScriptExecutor)d).ExecuteScript("return document.readyState").Equals("complete"));
                        if (driver.FindElements(By.XPath("//*[contains(@id,'DivLocationSelectTitle')]")).Count == 0)
                        {
                            Console.WriteLine("Corrupt folder/hyperlink  found  : " + hyperlink[i] + "\t" + FolderName + "\t" + path);
                        }

                        else
                        {
                            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
                            DownloadFileLink(driver, FolderName, false, Filepath, folderURLID);
                            WriteObjIDtoFile(hyperlink[i]);

                            //createIndex("Folder Name", FolderName, path, AllLinkItemSize[i].Text.ToString(), hyperlink[i]);
                            createIndex("Folder Name", FolderName.Substring(FolderName.LastIndexOf(("\\")) + 1), path, ItemSize[i], hyperlink[i]);
                        }



                    } /*else
                    {
                        Console.WriteLine("Folder not satisfying creation conditions : " + items[i] + "," + hyperlink[i]);
                        File.AppendAllText(CorruptFile, "Folder not satisfying creation conditions : " + " : " + items[i] + "\t," + hyperlink[i] + "\tCourrpt Folder" + Environment.NewLine); //Neha
                    }*/

                    if (hyperlink[i].ToString().Contains("1__125_"))
                    {
                        //Console.WriteLine("OMG OMG OMG" + items[i]);
                        driver.Navigate().GoToUrl(hyperlink[i]);
                        String FolderObjID = getObjID(hyperlink[i]);
                        wait.Until(d => ((IJavaScriptExecutor)d).ExecuteScript("return document.readyState").Equals("complete"));
                        try
                        {
                            var nextpage = driver.FindElement(By.XPath("//*[@id='PageNextImg']"));
                            if (nextpage.Enabled)
                            {

                                nextpage.Click();
                                Thread.Sleep(400);
                                Console.WriteLine("\n Click on Next pagefor debug only");
                                wait.Until(d => ((IJavaScriptExecutor)d).ExecuteScript("return document.readyState").Equals("complete"));
                                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMinutes(1);
                                Console.WriteLine("\n SleepTimeOver");
                                DownloadFileLink(driver, path, true, Filepath, FolderObjID);


                            }
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine(e);
                        }
                    }


                }
                //Here checking page if had multiple pages where document more than 125 item

            }

            catch (NoSuchElementException e)
            {
                Console.WriteLine("Done" + e);


            }

        }
        //Will go through the files in main folder, checks their creationtime against the dictionart, and if found will move the files to correct place
        public static void Version_unconfiremd(string filepath)
        {
            string[] fileEntries = Directory.GetFiles(filepath);
            foreach (string entry in fileEntries)
            {
                Console.WriteLine("/C wmic datafile where name='" + entry + "' get creationdate | findstr /brc:[0-9]");
                Process p = new Process();
                // Redirect the output stream of the child process.
                p.StartInfo.UseShellExecute = false;
                p.StartInfo.RedirectStandardOutput = true;
                p.StartInfo.FileName = @"C:\\Windows\\System32\\cmd.exe";
                p.StartInfo.CreateNoWindow = true;
                p.StartInfo.Arguments = "/C wmic datafile where name='" + entry.Replace("\\", "\\\\") + "' get creationdate | findstr /brc:[0-9]" + Environment.NewLine;
                p.OutputDataReceived += (sender, args) => Console.WriteLine("received output: {0}", args.Data);
                p.Start();

                p.WaitForExit();
                string output = p.StandardOutput.ReadToEnd();
                Console.WriteLine("Output:"+output);
                if (versionUnconfirmed.ContainsKey(output))
                {
                    Console.WriteLine(versionUnconfirmed[output]);
                    Directory.Move(entry, versionUnconfirmed[output]);
                }

            }
            foreach(string f in versionUnconfirmed.Keys)
            {
                Console.WriteLine(f);
            }
        }
        //get creation time of the file, and add to dictionary the correct place for the file, the key is the creationtime down to the microsecond
        private static void Get_Creationtime(string path,string source)
        {
            FileInfo f= new DirectoryInfo(source).GetFiles().OrderByDescending(o => o.LastWriteTime).FirstOrDefault();
            Console.WriteLine(f.Name);
            string name = f.FullName;
            if (f.Name.Contains(".crdownload")&&!f.Name.Contains("Unconfirmed "))
            {
                name=name.Replace(".crdownload", "");
            }
            string cmd = "/C wmic datafile where name='" + name.Replace("\\", "\\\\") + "' get creationdate | findstr /brc:[0-9]" + Environment.NewLine;
            Process p = new Process();
            // Redirect the output stream of the child process.
            p.StartInfo.UseShellExecute = false;
            p.StartInfo.RedirectStandardOutput = true;
            p.StartInfo.FileName = @"C:\\Windows\\System32\\cmd.exe";
            p.StartInfo.CreateNoWindow = true;
            Console.WriteLine(cmd);
            p.StartInfo.Arguments = cmd;
            p.OutputDataReceived += (sender, args) => Console.WriteLine("received output: {0}", args.Data);
            Console.WriteLine("process:"+p.Start());

            p.WaitForExit();
            string output = p.StandardOutput.ReadToEnd();
            Console.WriteLine(output);
            try
            {
                versionUnconfirmed.Add(output, path);
            }
            catch
            {
                Console.WriteLine("Problem with creationtime");
            }
            

        }
        //handles version control, Downloads each version, given a hyperlink to the file, saves it and tries to move it to correct location
        private static void Download_versions(string filehyperlink, IWebDriver driverObj, string sourcePathObj, string TargetPath)
        {
            Console.WriteLine(TargetPath);
            Actions actions = new Actions(driverObj);
            Console.WriteLine("downloadversion");
            if(!filehyperlink.Contains("objAction=browse")){
                if (!Directory.Exists(TargetPath + "//_MDMS Archive"))
                {
                    Directory.CreateDirectory(TargetPath + "/_MDMS Archive");
                }
                String fileVersionObjid = getObjID(filehyperlink);
                driverObj.Navigate().GoToUrl("https://mdms-ll.int.net.nokia.com/livelink/livelink.exe?func=ll&objId=" + fileVersionObjid + "&objAction=versions");
                IList<IWebElement> AllLinkVersion = driverObj.FindElements(By.ClassName("browseItemName"));//get links to the versions
                string filename = driverObj.FindElements(By.ClassName("pageTitleText"))[0].FindElement(By.TagName("h1")).Text;//get name of the file
                Console.WriteLine(filename);
                int downloaded = 0;
                Console.WriteLine(AllLinkVersion.Count);
                IJavaScriptExecutor js = (IJavaScriptExecutor)driverObj;
                int unconfirmed = 0;
                // loops through list of version links, downloads them and tries to save them to the correct place
                for (int i = 1; i < AllLinkVersion.Count; i++)
                {
                    actions.MoveToElement(AllLinkVersion[i]);
                    actions.Perform();
                    if (filename.Contains(".xml"))
                    {
                        AllLinkVersion[i].FindElements(By.TagName("a"))[1].Click();
                        IList<IWebElement> temp = driverObj.FindElements(By.ClassName("browseItemName"));
                        temp[i].FindElement(By.Id("menuItem_Download")).Click();
                    }
                    else
                    {
                        AllLinkVersion[i].FindElements(By.TagName("a"))[0].Click();
                    }
                    
                    IEnumerable<string> t = string.Join("", filename).Split(".");
                    String end = t.Last();
                    IEnumerable<string> b = t.SkipLast(1);
                    string path = TargetPath + "/_MDMS Archive/" + string.Join(".", b) + "_ver " + ((int)AllLinkVersion.Count - i).ToString() + "." + end; // construct a path, if it's longpath, create a longpath
                    if (path.Length > 255)
                    {
                        string lp = sourcePathObj + "/Log/longpath/" + DateTime.UtcNow.ToString("yyyyMMddHHmmss") + string.Join(".", b) + "_ver " + (AllLinkVersion.Count - i).ToString() + "." + end;
                        File.AppendAllText(Longpath, "move \"" + lp + "\" \"" + path + "\"" + Environment.NewLine);
                        Console.WriteLine("Longpath detected");
                        path = lp;
                    }
                    try
                    {
                        //checks if version is downloaded, and then moves it to correct place. If it has longpath, it will be moved to lonpath and line will be added to the file
                        Thread.Sleep(100);
                        FileInfo file = new DirectoryInfo(sourcePathObj).GetFiles().OrderByDescending(o => o.LastWriteTime).FirstOrDefault();
                        if (!file.Name.Contains("crdownload"))
                        {

                            if (path.Length > 255)
                            {
                                
                                string longp = sourcePathObj + "/Log/longpath/" + DateTime.UtcNow.ToString("yyyyMMddHHmmss") + string.Join(".", b) + "_ver " + (AllLinkVersion.Count - i).ToString() + "." + end;
                                Console.WriteLine(longp);
                                Directory.Move(file.FullName, longp);
                                File.AppendAllText(Longpath, "move \"" + longp + "\" \"" + path + "\"" + Environment.NewLine);
                                downloaded++;
                            }
                            else
                            {
                                Console.WriteLine(path);
                                Directory.Move(file.FullName, path);
                                downloaded++;
                            }
                        }
                        else
                        {
                            //if the file is downloading, add a thread, that will add the correct place of the file to the dictionary
                            
                            Console.WriteLine("unconfirmed");
                            string cmd="wmic datafile where name='" + file.FullName.Replace("\\", "\\\\") + "' get creationdate | findstr /brc:[0-9]" + Environment.NewLine;
                            new Thread(new ThreadStart(() => Get_Creationtime(path,sourcePathObj))).Start();
                            Thread.Sleep(50);
                            unconfirmed++;

                        }



                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex);
                    }
                }
                downloaded++;
                //if all versions were not downloaded, it will write to versionproblem file the problematic file, how many versions it has and how many werte downloaded, and where they are supposed to go
                if (downloaded < AllLinkVersion.Count)
                {
                    Console.WriteLine("all not downloaded");
                    File.AppendAllText(versionproblem, "https://mdms-ll.int.net.nokia.com/livelink/livelink.exe?func=ll&objId=" + fileVersionObjid + "&objAction=versions; " + "versions: " + AllLinkVersion.Count + ", Downloaded: " + downloaded.ToString() + ", location: " + TargetPath + "//_MDMS Archive/" + filename + Environment.NewLine);
                }

            }
            
        }


        private static void SkipUnsecureFile(IWebDriver SkipFiledriver, string Skipfilepath)
        {

            IList<IWebElement> AllLinkInFolder = SkipFiledriver.FindElements(By.ClassName("browseItemName"));  //Capture all available element
            for (int i = 0; i < AllLinkInFolder.Count; i++)
            {
                string Filehyperlink = AllLinkInFolder[i].FindElement(By.TagName("a")).GetAttribute("href");  //capture hyperlink of individual file to read objid to store in objfile
                String FileNametoCheck = AllLinkInFolder[i].Text.Trim();  //capture file name 
                if (!Filehyperlink.Contains("Action=browse") & (FileNametoCheck.EndsWith(".swf")))  // check if element is not a folder and file end with .swf
                {
                    Console.WriteLine("Skipped .swf File :" + FileNametoCheck);
                    WriteObjIDtoFile(Filehyperlink);
                    createIndex("SWF File", FileNametoCheck, Skipfilepath, "0", Filehyperlink);
                    //write objid in objID file which will help to skip file during download
                }

            }

        }
       /* private static string CheckSourceFile(string filenameObj, String DestinationFileObj, DateTime folderCreationTime)
        {
           
            Console.WriteLine("Original File duplicate check : " + filenameObj + " Destination is  : " + DestinationFileObj);
           if (File.Exists(filenameObj))
               {
                    DateTime FileCreationTime = getfileCreationTime(filenameObj);
                   if (folderCreationTime > FileCreationTime)
                    {
                        Console.WriteLine(" Duplicate file availabke skip and add to copyfailedbat : " + filenameObj);
                    }
                    else
                    {
                        Console.WriteLine("No File found which satisfy CFreation Time condition  :" + filenameObj);
                        
                    }

                }
                else
                {
                    Console.WriteLine("No File found or dowloaded which satisfy Creation Time condition  :" + filenameObj);ve
                    break;
                }
           

            return filenameObj;
        }
       */
        private static DateTime getfileCreationTime(string fileName)
        {
            DateTime CreationTime = File.GetCreationTime(fileName);
            return CreationTime;
        }

        private static DateTime getfolderCreationTime(string fileName)
        {
            DateTime CreationTime = Directory.GetCreationTime(fileName);
            return CreationTime;
        }
        private static string filenameCorerction(String FilenameObj)
        {
            string fileName = FilenameObj.Replace('%', '_').Replace('|', '_').Replace('/', '_').Replace("*", "").Replace("  ", "__").Replace("\"", "").Replace("'", "_").Replace('>', '_').Replace("?", "_").Replace('<', '_');
            fileName = fileName.Trim();
            if (fileName.Last() == '_' && !fileName.Contains(".html"))
            {
                fileName = fileName.Remove(fileName.Length - 1, 1);
            }
            if (fileName.Last() == '.')
            {
                fileName = fileName.Remove(fileName.Length - 1, 1);
            }

            return fileName;
        }

        private static void checkdownload(IWebDriver driver, String UserInput)
        {
            driver.Navigate().GoToUrl("chrome://downloads/");

            do
            {
                IJavaScriptExecutor jse2 = (IJavaScriptExecutor)driver;
                IWebElement root1 = (IWebElement)driver.FindElement(By.TagName("downloads-manager"));
                IWebElement downloadmanager = FindShadowRootElement(driver, root1);
                IWebElement FindShadowRootElement(IWebDriver Driver, IWebElement Selectors)
                {
                    IWebElement root = (IWebElement)((IJavaScriptExecutor)driver).ExecuteScript("return arguments[0].shadowRoot", Selectors);
                    return root;
                }

                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(60);

                if (downloadmanager.FindElements(By.TagName("downloads-item")).Count != 0)
                {
                    IWebElement root2 = (IWebElement)downloadmanager.FindElement(By.TagName("downloads-item"));
                    IWebElement item = FindShadowRootElement(driver, root2);
                    var button = item.FindElements(By.CssSelector("cr-button"));
                    for (int i = 0; i <= button.Count; i++)
                    {
                        if (button[i].GetAttribute("focus-type") == "save")
                        {
                            Console.WriteLine("Check All Unconfirmed download ");
                            UserInput = Console.ReadLine();
                            i = button.Count + 1;

                        }
                    }
                }
            } while (UserInput == "yes");

            driver.Navigate().Back();
        }
        private static string removeExtention(string sourcePathobj)
        {
            String extention = Path.GetExtension(sourcePathobj);
            if (extention == ".pdf")
            {
                sourcePathobj = sourcePathobj.Remove(sourcePathobj.Length - 4);
            }
            else if (extention == ".doc")
            {
                sourcePathobj = sourcePathobj.Remove(sourcePathobj.Length - 4);
            }
            else if (extention == ".docx")
            {
                sourcePathobj = sourcePathobj.Remove(sourcePathobj.Length - 5);
            }
            else if (extention == ".zip")
            {
                sourcePathobj = sourcePathobj.Remove(sourcePathobj.Length - 4);
            }
            else if (extention == ".xls")
            {
                sourcePathobj = sourcePathobj.Remove(sourcePathobj.Length - 4);
            }
            else if (extention == ".xlsx")
            {
                sourcePathobj = sourcePathobj.Remove(sourcePathobj.Length - 5);
            }

            return sourcePathobj;

        }
        private static Boolean readObjIDFile(String hyperlink, string filepath)
        {

            String[] line = File.ReadAllLines(filepath);
            int startIndex = hyperlink.IndexOf("&objId=");
            String objID = "not found";
            if ((hyperlink.Length > startIndex + 15) && (startIndex > 0))
            {
                objID = hyperlink.Substring(startIndex + 7, 8);

            }
            else
            {
                Console.WriteLine("No OBJ ID found in given URL -readobjID" + hyperlink);
            }


            return line.Contains(objID);
        }
        private static void WriteObjIDtoFile(string linkobj)
        {
            int startIndex = linkobj.IndexOf("&objId=");
            if ((linkobj.Length > startIndex + 15) && (startIndex > 0))
            {

                var objID = linkobj.Substring(startIndex + 7, 8);
                File.AppendAllText(objFilepath, objID + Environment.NewLine);
                Console.WriteLine(" ObjID has been written to file      :" + linkobj, "  :  " + objID);


            }
            else
            {
                Console.WriteLine("No OBJ ID found in given URL : writeobjid" + linkobj);
            }



        }

        private static String getObjID(string Elementlinkobj)
        {
            String getObjID = "UnknownOBJID";
            int startIndex = Elementlinkobj.IndexOf("&objId=");
            if ((Elementlinkobj.Length > startIndex + 15) && (startIndex > 0))
            {
                getObjID = Elementlinkobj.Substring(startIndex + 7, 8);
            }
            else
            {
                Console.WriteLine("No OBJ ID found in given URL : writeobjid" + Elementlinkobj);

            }

            return getObjID;
        }

        private static void createIndex(string typeOfFile, string Name, string path, string ObjectSize, string linkobj)
        {

            int startIndex = linkobj.IndexOf("&objId=");
            NumberOfline = File.ReadAllLines(IndexFormDMS).Length;

            if (NumberOfline > 50000)
            {
                File.AppendAllText(IndexFormDMS, "FolderOBJID,FileOBJID,Element Name,Size,AbsolutePath,URL");
                String IndexFormDMS2 = Filepath + "\\log\\" + "IndexFormDMS" + DateTime.Now.ToString("yyyymmddhhmmss") + ".csv";
                File.Move(IndexFormDMS, IndexFormDMS2);
                File.AppendAllText(IndexFormDMS, "Element ObjectID,FileOBJID,Element Name,Size,AbsolutePath,URL");

            }
            if ((linkobj.Length > startIndex + 15) && (startIndex > 0))
            {

                var objID = linkobj.Substring(startIndex + 7, 8);
                File.AppendAllText(IndexFormDMS, objID + "," + typeOfFile + "," + "\"" + Name + "\"" + "," + "\"" + path + "\"" + "," + ObjectSize + "," + linkobj + "," + Environment.NewLine);
                Console.WriteLine(" Detail has been mentioned in Index file  " + objID + "," + Name);

            }


            else
            {
                Console.WriteLine("No OBJ ID found in given URL : writeobjid" + linkobj);
            }

        }




        /*
         * @author Shajeda Akter Moni
         * @param htmlString
         * Email automation
         * Not in used at this moment
         */
        private static void Email(string htmlString)
        {
            try
            {
                MailMessage message = new MailMessage();
                SmtpClient smtp = new SmtpClient();
                message.From = new MailAddress("shajeda.moni@nokia.com");
                message.To.Add(new MailAddress("neha.1.goyal@nokia.com"));
                message.To.Add(new MailAddress("sari.kyllonen@nokia.com"));
                message.To.Add(new MailAddress("alok.khare@nokia.com"));
                message.Subject = "Test";
                message.IsBodyHtml = true; //to make message body as html  
                message.Body = htmlString;
                smtp.Port = 587;
                smtp.Host = "smtp.office365.com";
                smtp.EnableSsl = true;
                smtp.UseDefaultCredentials = false;
                smtp.Credentials = new System.Net.NetworkCredential("shajeda.moni@nokia.com", "Shaj234Pa");
                smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtp.Send(message);
                Console.WriteLine("Running the email");

            }
            catch (Exception)
            {
                Console.WriteLine("Not done Email automation");
            }
        }

    }
}

﻿using OpenQA.Selenium;
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;

namespace UzonePageObject
{
    public class UtilityHelper
    {
        public IWebDriver WebDriver;
        public WebDriverWait Wait { get { return new WebDriverWait(WebDriver, TimeSpan.FromSeconds(60)); } }

        //Takes a screenshot of an element. 
        public string TakeScreenshotOfElement(IWebDriver driver, IWebElement element, string name, string floor, int Id, bool person, string fileType = ".jpg")
        {
            string fileName;

            if (person)
            {
                fileName = Id + "-" + name.ToLower().Trim() + "-" + DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss") + "-" + floor.ToLower().Trim();
            }
            else
            {
                fileName = Id + "-seating_map-" + DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss") + "-" + floor.ToLower().Trim();
            }

            Byte[] byteArray = ((ITakesScreenshot)driver).GetScreenshot().AsByteArray;
            Bitmap screenshot = new Bitmap(new System.IO.MemoryStream(byteArray));
            System.Drawing.Rectangle croppedImage = new System.Drawing.Rectangle(element.Location.X, element.Location.Y, element.Size.Width, element.Size.Height);
            screenshot = screenshot.Clone(croppedImage, screenshot.PixelFormat);
            //string path = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), @"TestOutput\");

            string path = ReturnPathFolder(3, "TestOutput\\Screenshots");

            screenshot.Save(path + fileName + fileType, ImageFormat.Jpeg);

            return path + fileName + fileType;
        }

        public string ReturnPathFolder(int dirsBack = 3, string dirName = "TestOutput")
        {
            string properPath = "";
            string path = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            string[] dirs = path.Split('\\');
            int num = dirs.Length - dirsBack;

            for (int i = 0; i < num; i++)
            {
                properPath = properPath + dirs[i] + "\\";
            }

            return properPath + dirName + "\\";
        }

        public void EmptyDirectory(string path)
        {
            System.IO.DirectoryInfo di = new DirectoryInfo(path);

            foreach (FileInfo file in di.EnumerateFiles())
            {
                file.Delete();
            }
            foreach (DirectoryInfo dir in di.EnumerateDirectories())
            {
                dir.Delete(true);
            }
        }

        public void SetUpWebDriver()
        {
            //ChromeOptions options = new ChromeOptions();
            //options.AddArguments("--incognito");
            //options.AddArguments("--headless");
            //options.AddArguments("--user-agent=USFS");
            //options.AddArguments("--window-size=1920,1080");           

            //Instantiate new Chrome Driver and maximize the window
            //WebDriver = new ChromeDriver(options);
            WebDriver = new ChromeDriver();
            WebDriver.Manage().Window.Maximize();
            WebDriver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(25);
            WebDriver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(25);
        }

        public void TeardownWebDriver()
        {
            WebDriver.Quit();
        }

    }
}

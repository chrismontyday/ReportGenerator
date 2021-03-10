using OpenQA.Selenium;
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;

namespace UzonePageObject
{
    public class UtilityHelper
    {
        public IWebDriver WebDriver;

        //Takes a screenshot of an element. 
        public string TakeScreenshotOfElement(IWebDriver driver, IWebElement element, string name, int Id, bool person, string fileType = ".jpg")
        {
            string fileName;

            if (person)
            {
                fileName = Id + "-" + name.ToLower().Trim() + "-" + DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss");
            }
            else
            {
                fileName = Id + "-seating_map-" + DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss");
            }

            Byte[] byteArray = ((ITakesScreenshot)driver).GetScreenshot().AsByteArray;
            Bitmap screenshot = new Bitmap(new MemoryStream(byteArray));
            Rectangle croppedImage = new Rectangle(element.Location.X, element.Location.Y, element.Size.Width, element.Size.Height);
            screenshot = screenshot.Clone(croppedImage, screenshot.PixelFormat);
            string path = ReturnPathFolder(3, "TestOutput\\Screenshots");
            screenshot.Save(path + fileName + fileType, ImageFormat.Jpeg);            
            
            return path + fileName + fileType;
        }

        public string ReturnPathFolder(int dirsBack = 3, string dirName = "TestOutput")
        {
            string properPath = "";
            string path = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            string[] dirs = path.Split('\\');
            int num = dirs.Length - dirsBack;

            for (int i = 0; i < num; i++)
            {
                properPath = properPath + dirs[i] + "\\";
            }

            return properPath + dirName + "\\";
        }

        public byte[] GetImage(string filePath)
        {
            Image img = Image.FromFile(filePath);
            byte[] arr;
            using (MemoryStream ms = new MemoryStream())
            {
                img.Save(ms, ImageFormat.Jpeg);
                arr = ms.ToArray();
            }

            return arr;
        }

        public void ClearFolder(string filePath)
        {            
            DirectoryInfo di = new DirectoryInfo(filePath);

            foreach (FileInfo file in di.EnumerateFiles())
            {
                file.Delete();
            }
            foreach (DirectoryInfo dir in di.GetDirectories())
            {
                dir.Delete(true);
            }
        }

        public void KillExcel()
        {
            string strCmdText;
            strCmdText = "taskkill /f /im excel.exe";
            System.Diagnostics.Process.Start("CMD.exe", strCmdText);
        }

        public static string ConvertListToString(System.Collections.Generic.List<string> list)
        {
            string returnString = "\nBirthday & Anniversary Seating Map Report. Created on " + 
                DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss") + "\n";            

            for (int i = 0; i < list.Count; i++)
            {
                returnString = returnString + "\n" + list[i].ToString();
            }

            return returnString;
        }

    }
}

using System;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace UpdateResourcesLabels
{
   public partial class Form1 : Form
   {
      private const string basepath =
          //  @"C:\Users\Pavel.martiniuc\Desktop\Resx2Xls_source\Resx2Xls\UpdateResourcesLabels\bin\Debug\";
          @"D:\VsStudioProjects\LoxySoft\WFM-ProScheduler";
      private static StringBuilder missedKeys = new StringBuilder();
      private static XElement element;
      private static string resourceFile;
      public Form1()
      {
         InitializeComponent();
      }

      public static void UpdateResourceByKey(string pathToResources, string resourceFile, string key, string value)
      {

         if (element == null)
         {
            resourceFile = pathToResources + resourceFile;
            element = XElement.Load(resourceFile);
         }
         else
         {
            if (resourceFile != pathToResources + resourceFile)
            {
               resourceFile = pathToResources + resourceFile;
               element = XElement.Load(resourceFile);
            }
         }


         XElement val = element.Descendants("data").FirstOrDefault(el => el.Attribute("name").Value == key);
         if (val == null)
         {
            missedKeys.AppendLine(resourceFile + " | " + key + " | " + value);
            //XElement newElement = new XElement("data",
            //    new XAttribute("name", key),
            //    new XAttribute(XNamespace.Xml + "space", "preserve"),
            //    new XElement("value", value));
            // element.Add(newElement);

         }
         else
         {
            // val.Element("value").Value = value;
            if (val.Element("value").Value != value)
            {
               missedKeys.AppendLine(resourceFile + " | " + key + " | " + value);
            }
         }
         ///   element.Save(resourceFile);

      }

      private void button1_Click(object sender, EventArgs e)
      {
         // reqdXlRange(basepath + "ProScheduler_ResourcesForBrandReview_v3.xlsx", "A5", "E71", @"D:\VsStudioProjects\LoxySoft\WFM-ProScheduler\XLScheduler\ProScheduler\Resources\");
         // reqdXlRange(basepath + "ProScheduler_ResourcesForBrandReview_v3.xlsx", "A78", "E80", @"D:\VsStudioProjects\LoxySoft\WFM-ProScheduler\XLScheduler\ProScheduler.Wpf\Resources\");
         // reqdXlRange(basepath + "ProScheduler_ResourcesForBrandReview_v3.xlsx", "A85", "E103", @"D:\VsStudioProjects\LoxySoft\WFM-ProScheduler\ProSchedulerWeb\ProScheduler.WebUI\");
         UpdateGermanResources(
             @"D:\TASKS\REVIEWS\565_German_translation\translations.xlsx",
             "A3", "D5708", @"D:\PROJECTS\WFM-PROJECTS\WFM-ProScheduler");
      }


      public void reqdXlRange(string pathToXls, string startCoordinates, string endCoordinates, string pathToResources)
      {

         Excel.Application app = new Excel.Application();
         Excel.Workbook theWorkbook = app.Workbooks.Open(
             pathToXls, 0, true, 5,
             "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false,
             0, true);
         Excel.Sheets sheets = theWorkbook.Worksheets;
         Excel.Worksheet worksheet = (Excel.Worksheet)sheets.get_Item(1);

         Excel.Range range = worksheet.get_Range(startCoordinates, endCoordinates);
         System.Array myvalues = (System.Array)range.Cells.Value;
         Array strArray = myvalues;

         var arr = strArray as object[,];
         var len = arr.GetLength(0);

         for (int i = 1; i <= len; i++)
         {
            var resourceFile = (string)arr[i, 1];
            var key = (string)arr[i, 2];
            var oldValue = (string)arr[i, 3];
            var newValue = (string)arr[i, 4];
            UpdateResourceByKey(pathToResources, resourceFile, key, newValue);
         }

         app.Quit();

         GC.Collect();

      }

      public void UpdateGermanResources(string pathToXls, string startCoordinates, string endCoordinates, string pathToResources)
      {

         Excel.Application app = new Excel.Application();
         Excel.Workbook theWorkbook = app.Workbooks.Open(
             pathToXls, 0, true, 5,
             "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false,
             0, true);
         Excel.Sheets sheets = theWorkbook.Worksheets;
         Excel.Worksheet worksheet = (Excel.Worksheet)sheets.get_Item(1);

         Excel.Range range = worksheet.get_Range(startCoordinates, endCoordinates);
         System.Array myvalues = (System.Array)range.Cells.Value;
         Array strArray = myvalues;

         var arr = strArray as object[,];
         var len = arr.GetLength(0);
         missedKeys = new StringBuilder();
         for (int i = 1; i <= len; i++)
         {
            var resourceFile = (string)arr[i, 1];
            var key = (string)arr[i, 2];
            var oldValue = (string)arr[i, 3];
            var newValue = (string)arr[i, 4];
            UpdateResourceByKey(pathToResources, resourceFile, key, newValue);

            lbProcess.Text = (Convert.ToInt32((i / len)) * (int)100).ToString();
         }

         app.Quit();

         GC.Collect();
         rtbMissedKeys.Text = missedKeys.ToString();
      }

   }



}

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
      private static StringBuilder missedKeys = new StringBuilder();
      private static XElement ResourceElement;
      private static string resourceFile;
       private static int countDif = 0;
       private static int countNull = 0;
       private static int countAll = 2;
       private static int countTheSame = 0;

        public Form1()
      {
         InitializeComponent();
      }

      private void button1_Click(object sender, EventArgs e)
      {
            // reqdXlRange(basepath + "ProScheduler_ResourcesForBrandReview_v3.xlsx", "A5", "E71", @"D:\VsStudioProjects\LoxySoft\WFM-ProScheduler\XLScheduler\ProScheduler\Resources\");
            // reqdXlRange(basepath + "ProScheduler_ResourcesForBrandReview_v3.xlsx", "A78", "E80", @"D:\VsStudioProjects\LoxySoft\WFM-ProScheduler\XLScheduler\ProScheduler.Wpf\Resources\");
            // reqdXlRange(basepath + "ProScheduler_ResourcesForBrandReview_v3.xlsx", "A85", "E103", @"D:\VsStudioProjects\LoxySoft\WFM-ProScheduler\ProSchedulerWeb\ProScheduler.WebUI\");

          var pathToXls = @"D:\TASKS\REVIEWS\565_German_translation\translations.xlsx";
          var pathToResources = @"D:\PROJECTS\WFM-PROJECTS\WFM-ProScheduler";

          UpdateGermanResources(pathToXls, "A3", "D5708", pathToResources);
      }

      public void UpdateGermanResources(string pathToXls, string startCoordinates, string endCoordinates, string pathToResources)
      {

         Excel.Application app = new Excel.Application();
         Excel.Workbook theWorkbook = app.Workbooks.Open(
             pathToXls, 0, true, 5,
             "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true);
         Excel.Sheets sheets = theWorkbook.Worksheets;
         Excel.Worksheet worksheet = (Excel.Worksheet)sheets.Item[1];

         Excel.Range range = worksheet.Range[startCoordinates, endCoordinates];
         Array myvalues = (Array)range.Cells.Value;
         
         var arr = myvalues as object[,];
         
         missedKeys = new StringBuilder();

         for (int i = 1; i <= arr.GetLength(0); i++)
         {
            var resourceFile = (string)arr[i, 1];
            var key = (string)arr[i, 2];
            var oldValue = (string)arr[i, 3];
            var newValue = (string)arr[i, 4].ToString();
            
            UpdateResourceByKey(pathToResources, resourceFile, key, newValue);

            lbProcess.Text = (Convert.ToInt32((i / arr.GetLength(0))) * 100).ToString();
         }

         app.Quit();

         GC.Collect();

          missedKeys.AppendLine("countTheSame " + countTheSame + " countNull " + countNull + " | " + " countDif " + countDif + " | " + " countAll " + countAll);
          rtbMissedKeys.Text = missedKeys.ToString();
      }


       public static void UpdateResourceByKey(string pathToResources, string resourceFile, string key, string valueFromXls)
       {

           if (ResourceElement == null)
           {
               resourceFile = pathToResources + resourceFile;
               ResourceElement = XElement.Load(resourceFile);
               countAll += 1;
           }
           else
           {
               if (resourceFile != pathToResources + resourceFile)
               {
                   resourceFile = pathToResources + resourceFile;
                   ResourceElement = XElement.Load(resourceFile);
                   countAll += 1;
                }
           }


           XElement val = ResourceElement.Descendants("data").FirstOrDefault(el => el.Attribute("name")?.Value == key);

           if (val == null) // is not in resource file
           {
               missedKeys.AppendLine(countAll +" | NULL | " + resourceFile + " | " + key + " | " + valueFromXls);

                //XElement newElement = new XElement("data",
                //    new XAttribute("name", key),
                //    new XAttribute(XNamespace.Xml + "space", "preserve"),
                //    new XElement("value", value));
                // element.Add(newElement);
               countNull += 1;

           }
           else // diferent
           {
               if (val.Element("value")?.Value != valueFromXls)
               {
                   missedKeys.AppendLine(countAll + " | DIF | " + resourceFile + " | " + key + " | " + valueFromXls);
                   countDif += 1;
               }
               else
               {
                   // missedKeys.AppendLine(countAll + " | SAME | " + resourceFile + " | " + key + " | " + valueFromXls);
                   countTheSame += 1;
               }

           }
           ///   element.Save(resourceFile);
       }
    }
}

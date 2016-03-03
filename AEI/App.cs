using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AEI
{
    class ribbonUI : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication app)
        {
            string folderPath = @"C:\Program Files\AEI\AEI Tools for Revit 2016";
            string dll = Path.Combine(folderPath, "AEI.dll");

            string myRibbon = "AEI";
            app.CreateRibbonTab(myRibbon);

            RibbonPanel panelA = app.CreateRibbonPanel(myRibbon, "Electrical");
            /**
            PushButton btnOne = (PushButton)panelA.AddItem(new PushButtonData("Create Panel Sheets", "AEI", dll, "AEI.Command"));
            btnOne.ToolTip = "AEI";
            btnOne.LongDescription = "This command is great. It is really great. Really really wonderful!";

            PushButton btnTwo = (PushButton)panelA.AddItem(new PushButtonData("Create Sheets Set", "AEI", dll, "AEI.Command2"));
            btnOne.ToolTip = "AEI";
            btnOne.LongDescription = "This command is great. It is really great. Really really wonderful!";
            **/
            #region old code

            //PushButton btnTwo = (PushButton)panelA.AddItem(new PushButtonData("Two", "Two", dll, "BoostYourBIM.Two"));
            //btnTwo.LargeImage = new BitmapImage(new Uri("http://boostyourbim.files.wordpress.com/2013/08/note.png"));
            //btnTwo.Image = new BitmapImage(new Uri("http://boostyourbim.files.wordpress.com/2013/08/notesmall.png"));

            // Pull down button

            PulldownButtonData pullDownData = new PulldownButtonData("Panel Schedules", "Panel Schedules");

            //pullDownData.Image = new BitmapImage(new Uri(Path.Combine(folderPath, "image16.png"), UriKind.Absolute));
            //pullDownData.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "image32.png"), UriKind.Absolute));
            PulldownButton pullDownButton = panelA.AddItem(pullDownData) as PulldownButton;

            PushButton btnCreatePanelSheet = pullDownButton.AddPushButton(new PushButtonData("Create Panel Sheets", "Create Panel Sheet", dll, "AEI.Command"));
            PushButton btnCreateViewSet = pullDownButton.AddPushButton(new PushButtonData("PrintPDF", "Print Schedules to PDF", dll, "AEI.Command2"));
            //PushButton btnPrintPanetSheets = pullDownButton.AddPushButton(new PushButtonData("Print Panel Sheets", "Print Panel Sheets", dll, "AEI.Command3"));

            // Stacked list

            //PushButtonData dataHelp = new PushButtonData("Help", "Help", dll, "BoostYourBIM.help");
            //PushButtonData dataAbout = new PushButtonData("About", "About", dll, "BoostYourBIM.about");
            //PushButtonData dataFeedback = new PushButtonData("Feedback", "Feedback", dll, "BoostYourBIM.feedback");

            //IList<RibbonItem> stackedList = panelB.AddStackedItems(dataHelp, dataAbout, dataFeedback);

            //PushButton btnHelp = (PushButton)stackedList[0];
            //btnHelp.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "image32.png"), UriKind.Absolute));
            //btnHelp.Image = new BitmapImage(new Uri(Path.Combine(folderPath, "image16.png"), UriKind.Absolute));
            //btnHelp.ToolTip = "Click for our web-based help";

            //PushButton btnAbout = (PushButton)stackedList[1];
            //btnAbout.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "image32.png"), UriKind.Absolute));
            //btnAbout.Image = new BitmapImage(new Uri(Path.Combine(folderPath, "image16.png"), UriKind.Absolute));
            //btnAbout.ToolTip = "About these tools";

            //PushButton btnFeedback = (PushButton)stackedList[2];
            //btnFeedback.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "image32.png"), UriKind.Absolute));
            //btnFeedback.Image = new BitmapImage(new Uri(Path.Combine(folderPath, "image16.png"), UriKind.Absolute));
            //btnFeedback.ToolTip = "Click to email us with your feedback";
            #endregion

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication app)
        {
            return Result.Succeeded;
        }

    }

    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            Document doc = uiApp.ActiveUIDocument.Document;

            var count = 0;

            FilteredElementCollector filCol = new FilteredElementCollector(doc).OfClass(typeof(PanelScheduleView));

            foreach (var element in filCol)
            {

                using (Transaction t = new Transaction(doc, "Create a new ViewSheet"))
                {
                    t.Start();
                    try
                    {


                    // Create a sheet view
                    ViewSheet viewSheet = ViewSheet.Create(doc, ElementId.InvalidElementId);
                    if (null == viewSheet)
                    {
                        throw new Exception("Failed to create new ViewSheet.");
                    }

      

                    PanelScheduleSheetInstance.Create(doc, element.Id, viewSheet);
                    viewSheet.SheetNumber = "Panel " + count.ToString();
                    viewSheet.ViewName = element.Name;

                    count = count + 1;

                    t.Commit();
                }
                catch
                {
                    t.RollBack();
                }


            }

        }

                return Result.Succeeded;
        }
    }


    [Transaction(TransactionMode.Manual)]
    public class Command2 : IExternalCommand
    {

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            UIApplication uiApp = commandData.Application;
            Document doc = uiApp.ActiveUIDocument.Document;

            ViewSet myViewSet = new ViewSet();
            string match = "Panel";
            foreach (ViewSheet vs in new FilteredElementCollector(doc).OfClass(typeof(ViewSheet)).Cast<ViewSheet>()
                     .Where(q => q.SheetNumber.Contains(match)))
            {
                myViewSet.Insert(vs);
            }

            // get the PrintManger from the current document
            PrintManager printManager = doc.PrintManager;

            // set this PrintManager to use the "Selected Views/Sheets" option
            printManager.PrintRange = PrintRange.Select;

            // get the ViewSheetSetting which manages the view/sheet set information of current document
            ViewSheetSetting viewSheetSetting = printManager.ViewSheetSetting;

            // set the views in this ViewSheetSetting to the newly created ViewSet
            viewSheetSetting.CurrentViewSheetSet.Views = myViewSet;

            if (myViewSet.Size == 0)
            {
                TaskDialog.Show("Error", "No sheet numbers contain '" + match + "'.");
        
            }

            using (Transaction t = new Transaction(doc, "Create ViewSet"))
            {
                t.Start();
                string setName = "'" + match + "' Sheets";
                try
                {
                    // Save the current view sheet set to another view/sheet set with the specified name.
                    viewSheetSetting.SaveAs(setName);
                }
                // handle the exception that will occur if there is already a view/sheet set with this name
                catch (Autodesk.Revit.Exceptions.InvalidOperationException)
                {
                    TaskDialog.Show("Error", setName + " is already in use");
                    t.RollBack();
            
                }
                t.Commit();
            }
            TaskDialog.Show("View Set", myViewSet.Size + " sheets added to set");

            printManager.SelectNewPrintDriver("PDFCreator");
            printManager.PrintToFileName = @"C:\OutputFolder\" + Path.GetFileName(doc.PathName) + " Panel Schedules"  + ".";
            printManager.CombinedFile = true;
            printManager.SubmitPrint();

            return Result.Succeeded;

        }

    }


}

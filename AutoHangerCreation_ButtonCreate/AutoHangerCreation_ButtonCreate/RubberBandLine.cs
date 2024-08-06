using System;
using System.Windows.Forms;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.Exceptions;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Windows;
using Microsoft.VisualBasic.CompilerServices;
using Application = Autodesk.Revit.ApplicationServices.Application;

namespace AutoHangerCreation_ButtonCreate
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    internal class RubberBandLine : IExternalCommand
    {
        private IntPtr _revit_window;
        private List<ElementId> _added_element_ids;
        public RubberBandLine()
        {
            _added_element_ids = new List<ElementId>();
        }
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication application = commandData.Application;
            UIDocument activeUIDocument = application.ActiveUIDocument;
            Application application2 = application.Application;
            Document document = activeUIDocument.Document;
            _revit_window = application.MainWindowHandle;

            //找到RubberBand詳圖元素
            FilteredElementCollector val = new FilteredElementCollector(document);
            val.OfCategory((BuiltInCategory)(-2002000));//OST_DetailComponents = -2002000-->詳圖元素的品類
            val.OfClass(typeof(FamilySymbol));
            FamilySymbol val2 = (from FamilySymbol tag in (IEnumerable)new FilteredElementCollector(document).OfClass(typeof(FamilySymbol)).OfCategory((BuiltInCategory)(-2002000))
                                 where Operators.CompareString(((Element)tag).Name, "RubberBand", TextCompare: false) == 0
                                 select (tag)).First();
            _added_element_ids.Clear();
            application2.DocumentChanged += new EventHandler<DocumentChangedEventArgs>(OnDocumentChanged);
            while (true)
            {
                try
                {
                    UIDocument uidoc = commandData.Application.ActiveUIDocument;
                    Reference referPipe = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, "選擇測試");
                    activeUIDocument.PromptForFamilyInstancePlacement(val2);
                }
                //catch (OperationCanceledException val3)
                catch (Autodesk.Revit.Exceptions.OperationCanceledException val3)
                {
                    ProjectData.SetProjectError((Exception)val3);
                    //OperationCanceledException val4 = val3;
                    Autodesk.Revit.Exceptions.OperationCanceledException val4 = val3;
                    Debug.Print(((Exception)(object)val4).Message);
                    ProjectData.ClearProjectError();
                    break;
                }
            }
            application2.DocumentChanged -= new EventHandler<DocumentChangedEventArgs>(OnDocumentChanged);
            int count = _added_element_ids.Count;
            //MessageBox.Show($"有{_added_element_ids.Count}條參考線被創建");
            return Result.Succeeded;
        }
        private void OnDocumentChanged(object sender, DocumentChangedEventArgs e)
        {
            ICollection<ElementId> addedElementIds = e.GetAddedElementIds();
            int count = addedElementIds.Count;
            _added_element_ids.AddRange(addedElementIds);
            //if (_added_element_ids.Count >= 1 && ComponentManager.get_ApplicationWindow() != IntPtr.Zero)
            if (_added_element_ids.Count >= 1 && ComponentManager.ApplicationWindow != IntPtr.Zero)
            {
                //WindowsMessaging.PostWindowsMessage((int)ComponentManager.get_ApplicationWindow(), 256, 27, 0);
                //WindowsMessaging.PostWindowsMessage((int)ComponentManager.get_ApplicationWindow(), 256, 27, 0);

                //WindowsMessaging.PostWindowsMessage((int)ComponentManager.ApplicationWindow, 256, 27, 0);
                //WindowsMessaging.PostWindowsMessage((int)ComponentManager.ApplicationWindow, 256, 27, 0);
                //原程式有兩行"WindowsMessaging.PostWindowsMessage((int)ComponentManager.ApplicationWindow, 256, 27, 0)"
                //可以創建一次線段

                WindowsMessaging.PostWindowsMessage((int)ComponentManager.ApplicationWindow, 256, 27, 0);
                //如果修改成只有一行，可以創建多次線段

            }
        }
    }
}

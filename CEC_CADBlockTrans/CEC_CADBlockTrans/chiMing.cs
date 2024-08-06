#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Windows.Threading;

#endregion

namespace CEC_CADBlockTrans
{
    [Transaction(TransactionMode.Manual)]
    public class DeleteCADTypeMain : IExternalCommand
    {
        //宣告執行緒
        private Thread _uiThread;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            EventHandlerWithStringArg evStr = new EventHandlerWithStringArg();
            EventHandlerWithWpfArg evWpf = new EventHandlerWithWpfArg();
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            //紀錄使用者操作
            //LogWR logWR = new LogWR(doc, this.ToString());

            //啟動WPF事件
            //this.ShowFormSeparateThread(uiapp);
            return Result.Succeeded;
        }

        //public void ShowFormSeparateThread(UIApplication uiapp)
        //{
        //    // If we do not have a thread started or has been terminated start a new one
        //    if (!(_uiThread is null) && _uiThread.IsAlive) return;
        //    //EXTERNAL EVENTS WITH ARGUMENTS
        //    EventHandlerWithStringArg evStr = new EventHandlerWithStringArg();
        //    EventHandlerWithWpfArg evWpf = new EventHandlerWithWpfArg();
        //    // EventHandlerWithCorrectionArg evCor = new EventHandlerWithCorrectionArg();
        //    _uiThread = new Thread(() =>
        //    {
        //        SynchronizationContext.SetSynchronizationContext(
        //            new DispatcherSynchronizationContext(
        //                Dispatcher.CurrentDispatcher));
        //        // The dialog becomes the owner responsible for disposing the objects given to it.
        //        CADType window = new CADType(uiapp, evStr, evWpf);
        //        window.Closed += (s, e) => Dispatcher.CurrentDispatcher.InvokeShutdown();
        //        window.Show();
        //        Dispatcher.Run();
        //    });

        //    _uiThread.SetApartmentState(ApartmentState.STA);
        //    _uiThread.IsBackground = true;
        //    _uiThread.Start();
        //}
    }
}

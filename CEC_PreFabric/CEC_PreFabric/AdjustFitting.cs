#region Namespaces
using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.DB.Structure;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using System.IO;
using System.Threading;
#endregion

namespace CEC_PreFabric
{
    [Transaction(TransactionMode.Manual)]
    public class AdjustFitting : IExternalCommand
    {
#if RELEASE2019
        public static DisplayUnitType unitType = DisplayUnitType.DUT_MILLIMETERS;
#else
        public static ForgeTypeId unitType = UnitTypeId.Millimeters;
#endif
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Document doc = uidoc.Document;
            Autodesk.Revit.DB.View view = doc.ActiveView;

            ICollection<ElementId> ids = uidoc.Selection.GetElementIds();
            FamilySymbol targetSymbol = null;
            FilteredElementCollector coll = new FilteredElementCollector(doc).OfClass(typeof(Family));
            MessageBox.Show(coll.Count().ToString());
            string Famlist = "";
            int sucess = 0;
            //try
            //{
            //找到要用的彎頭
            foreach (Element e in coll)
            {
                string elemName = e.Name;
                Family fam = e as Family;
                Famlist += $"{elemName} , ";
                if (elemName == "GIP-彎頭(牙口)")
                {
                    targetSymbol = doc.GetElement(fam.GetFamilySymbolIds().First()) as FamilySymbol;
                }
            }
            //找到要置換的彎頭們
            List<Element> elemToChange = new List<Element>();

            FilteredElementCollector fittingColl = new FilteredElementCollector(doc, view.Id).OfClass(typeof(FamilyInstance)).OfCategory(BuiltInCategory.OST_PipeFitting);
            foreach (Element e in fittingColl)
            {
                FamilyInstance inst = e as FamilyInstance;
                FamilySymbol symbol = inst.Symbol;
                Family family = symbol.Family;
                if (family.Name == "彎頭 -牙口")
                {
                    elemToChange.Add(e);
                }
            }
            MessageBox.Show(elemToChange.Count().ToString());
            //using (TransactionGroup transGroup = new TransactionGroup(doc))
            //{
            //    transGroup.Start("彎頭更換");


                foreach (Element ele in elemToChange)
                {
                    using (Transaction trans = new Transaction(doc))
                    {
                        try
                        {
                            trans.Start("更換彎頭");
                            FailureHandlingOptions failOpt = trans.GetFailureHandlingOptions();

                            failOpt.SetFailuresPreprocessor(
                              new FittingWarningSwallower());

                            trans.SetFailureHandlingOptions(failOpt);

                            ElementId tempId = ele.ChangeTypeId(targetSymbol.Id);
                            if (tempId != null)
                            {
                                sucess += 1;
                            }
                            trans.Commit();
                        }
                        catch
                        {
                            MessageBox.Show("YO");
                            continue;
                        }

                    }
                }
                MessageBox.Show($"模型中共有「{elemToChange.Count()}」個彎頭需要置換，成功置換了「{sucess}」個");
                //foreach (ElementId id in ids)
                //{
                //    using (Transaction trans = new Transaction(doc))
                //    {
                //        trans.Start("更換彎頭");
                //        Element elem = doc.GetElement(id);
                //        elem.ChangeTypeId(targetSymbol.Id);
                //        trans.Commit();
                //    }
                //}
                //}
                //catch
                //{
                //    MessageBox.Show(Famlist);
                //    MessageBox.Show("彎頭更換失敗");
                //}
            //    transGroup.Assimilate();
            //}

            return Result.Succeeded;
        }
    }
    public class FittingWarningSwallower : IFailuresPreprocessor
    {
        public FailureProcessingResult PreprocessFailures(
          FailuresAccessor a)
        {
            // inside event handler, get all warnings
            List<FailureDefinitionId> exceptionID = new List<FailureDefinitionId>()
            {
               BuiltInFailures.AutoRouteFailures.FittingParamsCauseGeometryToFailWarning,
               BuiltInFailures.AutoRouteFailures.FittingParamsCauseGeometryToFailError,
               BuiltInFailures.AutoRouteFailures.NoFittingInRoutingError,
               BuiltInFailures.AutoRouteFailures.SegmentDrawnFromWrongSideError,
               BuiltInFailures.AutoRouteFailures.DesignOptionConflictWarning,
               BuiltInFailures.AutoRouteFailures.SolutionIsOppositeToConnectionError,
               BuiltInFailures.AutoRouteFailures.SolutionIsOppositeToConnectionWarning,
               BuiltInFailures.AutoRouteFailures.SegmentDrawnFromWrongSideWarning,
               BuiltInFailures.AutoRouteFailures.AttemptToConnectNonSlopingElementToSlopedPipeError,
               BuiltInFailures.AutoRouteFailures.AttemptToConnectNonSlopingElementToSlopedPipeWarning,
               BuiltInFailures.AutoRouteFailures.DesignOptionConflictError,
               BuiltInFailures.AutoRouteFailures.DesignOptionConflictWarning,
               BuiltInFailures.AutoRouteFailures.DistributionTypeInvalid,
               BuiltInFailures.AutoRouteFailures.ElementHasFlowCalculation,
               BuiltInFailures.AutoRouteFailures.ElementHasOpenConnection,
               BuiltInFailures.AutoRouteFailures.FailedDowngradeCrossToElbow,
               BuiltInFailures.AutoRouteFailures.FailedDowngradeCrossToTee,
               BuiltInFailures.AutoRouteFailures.FailedDowngradeTeeToCross,
               BuiltInFailures.AutoRouteFailures.HighlightedElementDisconnected,
               BuiltInFailures.AutoRouteFailures.MismatchedSystemClassificationHasSystemWarning,
               BuiltInFailures.AutoRouteFailures.MismatchedSystemClassificationNoSystemWarning,
               BuiltInFailures.AutoRouteFailures.NoAutoRouteSolutionFoundError,
               BuiltInFailures.AutoRouteFailures.NoAutoRouteSolutionFoundWarning,
               BuiltInFailures.AutoRouteFailures.NoFittingInRoutingError,
               BuiltInFailures.AutoRouteFailures.NoValidSolutionAtSelectedPoint,
               BuiltInFailures.AutoRouteFailures.SegmentDrawnFromWrongSideWarning,
               BuiltInFailures.AutoRouteFailures.SpecifiedAnglesNotCompatibleWithLayoutError,
               BuiltInFailures.AutoRouteFailures.SpecifiedAnglesNotCompatibleWithLayoutWarning,
               BuiltInFailures.AutoRouteFailures.UnableToMakeConnectionInDirection,
               BuiltInFailures.AutoRouteFailures.ZeroLengthSegmentError,
               BuiltInFailures.AutoRouteFailures.ZeroLengthSegmentWarning
            };

            IList<FailureMessageAccessor> failures
              = a.GetFailureMessages();

            foreach (FailureMessageAccessor f in failures)
            {
                // check failure definition ids 
                // against ones to dismiss:

                FailureDefinitionId id
                  = f.GetFailureDefinitionId();
                MessageBox.Show(id.ToString());

                //if (exceptionID.Contains(id))
                //  //  BuiltInFailures.AutoRouteFailures.FittingParamsCauseGeometryToFailWarning
                //  //== id)
                //{
                    a.DeleteAllWarnings();
                    //a.DeleteWarning(f);
                //}
            }
            return FailureProcessingResult.Continue;
        }
    }
}

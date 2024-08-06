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
    //【程式架構】
    //1.先讓使用者可以選擇elements
    //2.創建ISOmetric
    //3.利用選擇的elements來改變ISO的orientation

    [Transaction(TransactionMode.Manual)]
    public class CreateISO : IExternalCommand
    {
#if RELEASE2019
        public static DisplayUnitType unitType = DisplayUnitType.DUT_MILLIMETERS;
#else
        public static ForgeTypeId unitType = UnitTypeId.Millimeters;
#endif
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Counter.count += 1;
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Document doc = uidoc.Document;
            try
            {
                //取得試圖樣板測試
                PreFabricUI ui = new PreFabricUI();
                FilteredElementCollector coll = new FilteredElementCollector(doc);
                List<Autodesk.Revit.DB.View> templateList = new List<Autodesk.Revit.DB.View>();
                coll.OfCategory(BuiltInCategory.OST_Views);
                foreach (Element v in coll)
                {
                    Autodesk.Revit.DB.View tempView = v as Autodesk.Revit.DB.View;
                    if (tempView.IsTemplate)
                    {
                        templateList.Add(tempView);
                    }
                }
                ui.startingNumTextBox.IsReadOnly = true;
                ui.viewTemplateComboBox.ItemsSource = templateList;
                ui.ShowDialog();
                //利用DialogResult來判斷UI是否成功執行
                if (ui.DialogResult == false)
                {
                    return Result.Failed;
                }
                string viewName = ui.viewNameTextBox.Text;
                string levelName = ui.levelName.Text;
                string regionName = ui.regionName.Text;
                Autodesk.Revit.DB.View selectedView = ui.viewTemplateComboBox.SelectedItem as Autodesk.Revit.DB.View;
                List<string> checkList = new List<string>() { viewName, levelName, regionName };
                bool cancel = false;
                if (selectedView == null)
                {
                    cancel = true;
                }
                else
                {
                    foreach (string str in checkList)
                    {
                        if (str == "")
                        {
                            cancel = true;
                            break;
                        }
                    }
                }
                if (cancel == true)
                {
                    MessageBox.Show("視窗上的內容不可為空，\n請檢查「視圖名稱」、「視圖樣版」、「裁切料號」參數是否都有正確填入!");
                    return Result.Failed;
                }
                //拿到管元件-先測試本地的
                ISelectionFilter pipeFilter = new PipeSelectionFilter(doc);
                ISelectionFilter linkedPipeFilter = new linkedPipeSelectionFilter(doc);
                IList<Reference> pickPipeRefs = uidoc.Selection.PickObjects(ObjectType.Element, pipeFilter, "請選擇要製作軸測圖的管元件");
                List<Element> pickPipes = new List<Element>();

                foreach (Reference refer in pickPipeRefs)
                {
                    Element tempPipe = doc.GetElement(refer);
                    pickPipes.Add(tempPipe);
                }
                using (TransactionGroup transGroup = new TransactionGroup(doc))
                {
                    transGroup.Start("產生預製ISO圖");
                    //step1－產生ISO圖並把ISO圖變為當前視圖
                    Autodesk.Revit.DB.View isoView = createNewViewAndZoom(commandData);

                    //step2 - 變更視圖剪裁範圍 (看是否需要設定剖面框)
                    using (Transaction trans = new Transaction(doc))
                    {
                        trans.Start("變更視圖剪裁盒");
                        BoundingBoxXYZ bounding = isoView.CropBox;
                        Transform transform = bounding.Transform;
                        Transform transInverse = transform.Inverse;
                        List<XYZ> points = new List<XYZ>();
                        XYZ ptWork = null;
                        //將世界座標轉換成視圖座標，透過TransformInverse的逆變來完成
                        foreach (Reference r in pickPipeRefs)
                        {
                            BoundingBoxXYZ bb = doc.GetElement(r).get_BoundingBox(null);
                            ptWork = transInverse.OfPoint(bb.Min);
                            points.Add(ptWork);
                            ptWork = transInverse.OfPoint(bb.Max);
                            points.Add(ptWork);
                        }
                        double adjust = 500;
                        double adjustOffset = UnitUtils.ConvertToInternalUnits(adjust, unitType);
                        BoundingBoxXYZ sb = new BoundingBoxXYZ();
                        sb.Min = new XYZ(points.Min(p => p.X - adjustOffset),
                                          points.Min(p => p.Y - adjustOffset),
                                          points.Min(p => p.Z - adjustOffset));
                        sb.Max = new XYZ(points.Max(p => p.X + adjustOffset),
                                       points.Max(p => p.Y + adjustOffset),
                                       points.Max(p => p.Z + adjustOffset));
                        isoView.CropBox = sb;
                        isoView.CropBoxActive = false;
                        isoView.CropBoxVisible = false;
                        isoView.DetailLevel = ViewDetailLevel.Fine;
                        isoView.DisplayStyle = DisplayStyle.FlatColors;

                        //變更SectionBox的狀態
                        //isoView3D.GetSectionBox();
                        List<XYZ> boundingCorners = new List<XYZ>();
                        List<double> boundingX = new List<double>();
                        List<double> boundingY = new List<double>();
                        List<double> boundingZ = new List<double>();
                        foreach (Element e in pickPipes)
                        {
                            BoundingBoxXYZ tempBox = e.get_BoundingBox(null);
                            XYZ maxCorner = tempBox.Max;
                            XYZ minCorner = tempBox.Min;
                            boundingCorners.Add(maxCorner);
                            boundingCorners.Add(minCorner);
                            //拆解XYZ值之後加入
                            boundingX.Add(maxCorner.X);
                            boundingX.Add(minCorner.X);
                            boundingY.Add(maxCorner.Y);
                            boundingY.Add(minCorner.Y);
                            boundingZ.Add(maxCorner.Z);
                            boundingZ.Add(minCorner.Z);
                        }
                        boundingX = boundingX.OrderByDescending(x => x).ToList();
                        boundingY = boundingY.OrderByDescending(y => y).ToList();
                        boundingZ = boundingZ.OrderByDescending(z => z).ToList();
                        XYZ maxBB = new XYZ(boundingX.First(), boundingY.First(), boundingZ.First());
                        XYZ minBB = new XYZ(boundingX.Last(), boundingY.Last(), boundingZ.Last());

                        boundingCorners = boundingCorners.OrderByDescending(x => x.X).ThenByDescending(x => x.X).ThenByDescending(x => x.Y).ToList();
                        BoundingBoxXYZ newBox = new BoundingBoxXYZ();
                        newBox.Max = maxBB;
                        newBox.Min = minBB;
                        //轉成3D視圖後才可以鎖定視圖
                        View3D isoView3D = isoView as View3D;
                        isoView3D.SaveOrientationAndLock();
                        isoView3D.IsSectionBoxActive = true;
                        isoView3D.SetSectionBox(newBox);
                        isoView3D.Name = viewName;
                        isoView3D.ViewTemplateId = selectedView.Id;
                        trans.Commit();
                    }

                    //step3 - 針對特殊元素加入共用參數，並加以分組寫入數字
                    #region 找到要載入共用參數的品類
                    //1.要先確認這個binding是否存在
                    //2.確認想寫入的品類裡有沒有這個參數，如果有則去除這個品類
                    //3.將剩下的品類寫入既有binding
                    List<string> paraName1 = new List<string>() {"【預組】系統別", "【預組】樓層", "【預組】區域", "【預組】編號" };
                    //List<string> paraName = new List<string>() { "管料裁切編號", "裁切料號" };
                    string checkString = "";
                    //foreach (string st in paraName)
                    foreach (string st in paraName1)
                    {
                        List<Category> defaultCateList = new List<Category>()
                {
                    Category.GetCategory(doc,BuiltInCategory.OST_PipeCurves),
                    Category.GetCategory(doc,BuiltInCategory.OST_DuctCurves),
                    Category.GetCategory(doc,BuiltInCategory.OST_Conduit)
                };
                        CategorySet catSet = app.Create.NewCategorySet();
                        foreach (Element e in pickPipes)
                        {
                            Category tempCate = e.Category;
                            if (!catSet.Contains(tempCate))
                            {
                                catSet.Insert(tempCate);
                            }
                        }
                        BindingMap bm = doc.ParameterBindings;
                        DefinitionBindingMapIterator itor = bm.ForwardIterator();
                        itor.Reset();
                        Definition d = null;
                        ElementBinding elemBind = null;
                        //如果現在的專案中已經載入該參數欄位，則不需重新載入
                        while (itor.MoveNext())
                        {
                            d = itor.Key;
                            if (d.Name == st)
                            {
                                elemBind = (ElementBinding)itor.Current;
                                break;
                            }
                        }
                        //如果該共用參數已經載入成為專案參數，重新加入binding
                        if (d.Name == st && catSet.Size > 0)
                        {
                            using (Transaction tx = new Transaction(doc, "Add Binding"))
                            {
                                tx.Start();
                                InstanceBinding ib = doc.Application.Create.NewInstanceBinding(catSet);
                                bool result = doc.ParameterBindings.ReInsert(d, ib, BuiltInParameterGroup.PG_SEGMENTS_FITTINGS);
                                tx.Commit();
                            }
                        }
                        //如果該專案參數還沒被載入，則載入之
                        else if (d.Name != st)
                        {
                            //MessageBox.Show($"專案尚未載入「 {spName}」 參數，將自動載入");
                            checkString += $"專案尚未載入「 {st}」 參數，將自動載入\n";
                            //MessageBox.Show($"專案尚未載入「 {st}」 參數，將自動載入");
                            var infoPath = @"Dropbox\info.json";
                            var jsonPath = Path.Combine(Environment.GetEnvironmentVariable("LocalAppData"), infoPath);
                            if (!File.Exists(jsonPath)) jsonPath = Path.Combine(Environment.GetEnvironmentVariable("AppData"), infoPath);
                            if (!File.Exists(jsonPath)) throw new Exception("請安裝並登入Dropbox桌面應用程式!");
                            var dropboxPath = File.ReadAllText(jsonPath).Split('\"')[5];
                            var spFilePath = dropboxPath + @"\BIM-Share\BIM共用參數.txt";
                            app.SharedParametersFilename = spFilePath;
                            DefinitionFile spFile = app.OpenSharedParameterFile();
                            ExternalDefinition targetDefinition = null;
                            foreach (DefinitionGroup dG in spFile.Groups)
                            {
                                if (dG.Name == "機電_預組")
                                {
                                    foreach (ExternalDefinition def in dG.Definitions)
                                    {
                                        if (def.Name == st) targetDefinition = def;
                                    }
                                }
                            }
                            //在此之前要建立一個審核該參數是否已經被載入的機制，如果已被載入則不載入
                            if (targetDefinition != null)
                            {
                                using (Transaction trans = new Transaction(doc))
                                {
                                    trans.Start("載入共用參數");
                                    InstanceBinding newIB = app.Create.NewInstanceBinding(catSet);
                                    doc.ParameterBindings.Insert(targetDefinition, newIB, BuiltInParameterGroup.PG_SEGMENTS_FITTINGS);
                                    trans.Commit();
                                }
                            }
                            else if (targetDefinition == null)
                            {
                                MessageBox.Show($"共用參數中沒有找到 {st} 參數");
                            }
                        }
                    }
                    //MessageBox.Show(checkString);
                    #endregion
                    //step4 - 針對視圖中的管材加上tag，分組並上入編號-->分組的寫法該怎麼寫還待思考
                    string filterName = "";
                    int keyToSet = 1;
                    List<Element> pipeListToCheck = pickPipes; //新增另外一個List去做交叉比對，利用他來對照pickPipes中的元件，以及刪除已經配對到的
                    List<Element> alreadySetList = new List<Element>(); //新增一個List去蒐集已經編號過的
                    Element tempElem = null;
                    string outPut = "";
                    foreach (Element e in pickPipes)
                    {
                        List<Element> sameList = new List<Element>();
                        tempElem = e;
                        //比較的基準-->材料(名稱)、管徑(大小)、長度
                        string elemName = e.Name;
                        string elemSize = getPipeDiameter(e);
                        double elemLength = Math.Round(e.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble(), 2);
                        //string elemFabName = e.LookupParameter(spName).AsString();
                        //if (elemFabName != "")
                        //{
                        //    MessageBox.Show("選中管件中已寫入裁切編號，請清除編號後再重新寫入");
                        //    return Result.Failed;
                        //}
                        foreach (Element ee in pickPipes)
                        {
                            double tempLength = Math.Round(ee.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble(), 2);
                            if (elemName == ee.Name && elemSize == getPipeDiameter(ee) && !alreadySetList.Contains(ee) && elemLength == tempLength/*&& ee.LookupParameter(spName).AsString() == ""*/)
                            {
                                sameList.Add(ee);
                                alreadySetList.Add(ee);
                            }
                        }
                        if (sameList.Count() > 0)
                        {
                            outPut += $"與名為{elemName}，且大小同為{elemSize}的管材共有{sameList.Count()}個\n";
                            using (Transaction trans = new Transaction(doc))
                            {
                                trans.Start("寫入裁切編號");
                                foreach (Element p in sameList)
                                {
                                    //Parameter fabNum = p.LookupParameter(paraName[0]);
                                    //Parameter fabFullName = p.LookupParameter(paraName[1]);
                                    Parameter pipeSystem = p.LookupParameter(paraName1[0]);
                                    Parameter pipeLevel = p.LookupParameter(paraName1[1]);
                                    Parameter pipeRegion = p.LookupParameter(paraName1[2]);
                                    Parameter pipeNum = p.LookupParameter(paraName1[3]);

                                    string systemName = p.get_Parameter(BuiltInParameter.RBS_DUCT_PIPE_SYSTEM_ABBREVIATION_PARAM).AsString();
                                    //string numToSet = systemName + "-" + levelName + "-" + regionName + "-" + keyToSet.ToString();
                                    filterName = systemName + "-" + levelName + "-" + regionName;
                                    //fabNum.Set(keyToSet.ToString());
                                    //fabFullName.Set(numToSet);
                                    pipeSystem.Set(systemName);
                                    pipeLevel.Set(levelName);
                                    pipeRegion.Set(regionName);
                                    pipeNum.Set(keyToSet.ToString());

                                }
                                keyToSet += 1;
                                trans.Commit();
                            }
                        }
                    }
                    //針對每個元件放上多重品類標籤
                    Element multiCateTag = findMultiCateTag(doc);
                    if (multiCateTag == null) MessageBox.Show("專案尚未載入預組視圖專用標籤");
                    using (Transaction trans = new Transaction(doc))
                    {
                        trans.Start("放置多重品類標籤");
                        foreach (Element e in pickPipes)
                        {
                            Reference elemRefer = new Reference(e);
                            MEPCurve mepCrv = e as MEPCurve;
                            LocationCurve pipeLocate = mepCrv.Location as LocationCurve;
                            Curve pipeCrv = pipeLocate.Curve;
                            XYZ middlePt = pipeCrv.Evaluate(0.5, true);
                            IndependentTag fabricTag = IndependentTag.Create(doc, multiCateTag.Id, isoView.Id, elemRefer, true, TagOrientation.Horizontal, middlePt);
                            XYZ tagHead = fabricTag.TagHeadPosition;
                            //設定隨機產生一定offset值的方法
                            Line tempLine = Line.CreateBound(middlePt, tagHead);
                            XYZ targetPt = tempLine.Evaluate(0.3, true);
                            fabricTag.TagHeadPosition = targetPt;
                        }
                        trans.Commit();
                    }
                    //MessageBox.Show(outPut);

                    //step5 創造管段明細表
                    //List<string> scheduleParas = new List<string>()
                    //{
                    //    "裁切料號","系統類型","管料裁切編號","大小","長度","數量"
                    //};
                    List<string> tempList = new List<string>() { "大小", "長度", "數量" };
                    List<string> scheduleParas = new List<string>();
                    scheduleParas = paraName1.Union(tempList).ToList();
                    using (Transaction trans = new Transaction(doc))
                    {
                        trans.Start("創建管明細表");
                        ElementId pipeCateId = new ElementId(BuiltInCategory.OST_PipeCurves);
                        ViewSchedule schedule = ViewSchedule.CreateSchedule(doc, pipeCateId);
                        schedule.Name = viewName + "管料裁切明細表";
                        //ScheduleFilter numFilter = null;
                        ScheduleFilter systemFilter = null;
                        ScheduleFilter levelFilter = null;
                        ScheduleFilter regionFilter = null;
                        ScheduleSortGroupField numberSorting = null;
                        foreach (string str in scheduleParas)
                        {
                            foreach (SchedulableField sf in schedule.Definition.GetSchedulableFields())
                            {
                                if (sf.GetName(doc) == str)
                                {
                                    {
                                        ScheduleField scheduleField = schedule.Definition.AddField(sf);
                                        ////1.以系統進行篩選
                                        //if (sf.GetName(doc) == paraName1[0])
                                        //{
                                        //    systemFilter = new ScheduleFilter(scheduleField.FieldId,ScheduleFilterType.Equal.systemName)
                                        //}
                                        //2.以樓層進行篩選
                                       if (sf.GetName(doc) == paraName1[1])
                                        {
                                            levelFilter = new ScheduleFilter(scheduleField.FieldId, ScheduleFilterType.Equal, levelName);
                                        }
                                        //3.以區域進行篩選
                                        else if (sf.GetName(doc) == paraName1[2])
                                        {
                                            regionFilter = new ScheduleFilter(scheduleField.FieldId, ScheduleFilterType.Equal, regionName);
                                        }
                                        //4.以編號進行排序
                                        else if (sf.GetName(doc) == paraName1[3])
                                        {
                                            numberSorting = new ScheduleSortGroupField(scheduleField.FieldId);
                                        }
                                        #region 以前只匯入裁切料號與管料裁切編號的舊寫法
                                        ////以裁切料號進行明細表元件的篩選
                                        //if (sf.GetName(doc) == "裁切料號")
                                        //{
                                        //    numFilter = new ScheduleFilter(scheduleField.FieldId, ScheduleFilterType.Contains, filterName);
                                        //}
                                        ////以管料裁切編號進行排序
                                        //if (sf.GetName(doc) == "管料裁切編號")
                                        //{
                                        //    numberSorting = new ScheduleSortGroupField(scheduleField.FieldId);
                                        //}
                                        #endregion
                                    }
                                }
                            }
                            #region 添加欄位的舊寫法
                            //foreach (SchedulableField sf in schedule.Definition.GetSchedulableFields())
                            //{
                            //    if (scheduleParas.Contains(sf.GetName(doc)))
                            //    {
                            //        ScheduleField scheduleField = schedule.Definition.AddField(sf);
                            //        if (sf.GetName(doc) == "裁切料號")
                            //        {
                            //            numFilter = new ScheduleFilter(scheduleField.FieldId, ScheduleFilterType.Contains, filterName);
                            //        }
                            //    }
                            //}
                            #endregion
                            //目前尚缺減去詳細列舉每個實體的機制&針對明細表攔位的排序
                        }
                        //schedule.Definition.AddFilter(numFilter);
                        schedule.Definition.AddFilter(levelFilter);
                        schedule.Definition.AddFilter(regionFilter);
                        schedule.Definition.AddSortGroupField(numberSorting);
                        schedule.Definition.IsItemized = false;
                        trans.Commit();
                    }
                    transGroup.Assimilate();
                }
            }
            catch
            {
                MessageBox.Show("執行失敗");
                return Result.Failed;
            }
            return Result.Succeeded;
        }
        //製作一個方法找到多重品類標籤
        public Element findMultiCateTag(Document doc)
        {
            string tagetName = "M_裁切編號標籤";
            Element targetElement = null;
            FilteredElementCollector coll = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_MultiCategoryTags).WhereElementIsElementType();
            foreach (Element e in coll)
            {
                if (e.Name == tagetName) targetElement = e;
            }
            return targetElement;
        }
        public string getPipeDiameter(Element elem)
        {
            Document _doc = elem.Document;
            Category pipe = Category.GetCategory(_doc, BuiltInCategory.OST_PipeCurves);
            Category duct = Category.GetCategory(_doc, BuiltInCategory.OST_DuctCurves);
            Category conduit = Category.GetCategory(_doc, BuiltInCategory.OST_Conduit);
            Category tray = Category.GetCategory(_doc, BuiltInCategory.OST_CableTray);
            double targetValue = 0.0;
            string targetValueStr = "";
            if (elem.Category.Id == pipe.Id)
            {
                targetValueStr = elem.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE).AsString();
            }
            else if (elem.Category.Id == duct.Id)
            {
                targetValueStr = elem.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE).AsString();
            }
            else if (elem.Category.Id == conduit.Id)
            {
                targetValueStr = elem.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE).AsString();
            }
            else if (elem.Category.Id == tray.Id)
            {
                targetValueStr = elem.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE).AsString();
            }
            return targetValueStr;
        }
        public Autodesk.Revit.DB.View createNewViewAndZoom(ExternalCommandData commandData)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uiapp.ActiveUIDocument.Document;
            ViewFamilyType vft3d = new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>().FirstOrDefault(q => q.ViewFamily == ViewFamily.ThreeDimensional);
            Autodesk.Revit.DB.View view = null;
            using (Transaction t = new Transaction(doc, "make views"))
            {
                t.Start();
                view = View3D.CreateIsometric(doc, vft3d.Id);
                t.Commit();
            }
            zoom(uidoc, view);
            return view;
        }
        private void zoom(UIDocument uidoc, Autodesk.Revit.DB.View view)
        {
            // YOU NEED THESE TWO LINES OR THE ZOOM WILL NOT HAPPEN!
            uidoc.ActiveView = view;
            uidoc.RefreshActiveView();

            UIView uiview = uidoc.GetOpenUIViews().Cast<UIView>().FirstOrDefault(q => q.ViewId == view.Id);
            uiview.Zoom(5);
        }
        public void CropViewBySelection(Document doc, IList<Reference> pickPipeRefs, Autodesk.Revit.DB.View v)
        {
            //Autodesk.Revit.DB.View activeView = doc.ActiveView;
            List<XYZ> points = new List<XYZ>();
            foreach (Reference r in pickPipeRefs)
            {
                BoundingBoxXYZ bb = doc.GetElement(r).get_BoundingBox(v);
                points.Add(bb.Min);
                points.Add(bb.Max);
            }
            BoundingBoxXYZ sb = new BoundingBoxXYZ();
            sb.Min = new XYZ(points.Min(p => p.X),
                              points.Min(p => p.Y),
                              points.Min(p => p.Z));
            sb.Max = new XYZ(points.Max(p => p.X),
                           points.Max(p => p.Y),
                           points.Max(p => p.Z));
            try
            {
                using (Transaction t = new Transaction(doc, "Crop View By Selection"))
                {
                    t.Start();
                    v.CropBoxActive = true;
                    v.CropBox = sb;
                    t.Commit();
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Revit", ex.Message);
                return;
            }
        }
    }
    public class PipeSelectionFilter : ISelectionFilter
    {
        private Document _doc;
        public PipeSelectionFilter(Document doc)
        {
            this._doc = doc;
        }
        public bool AllowElement(Element element)
        {
            Category pipe = Category.GetCategory(_doc, BuiltInCategory.OST_PipeCurves);
            Category duct = Category.GetCategory(_doc, BuiltInCategory.OST_DuctCurves);
            Category conduit = Category.GetCategory(_doc, BuiltInCategory.OST_Conduit);
            Category tray = Category.GetCategory(_doc, BuiltInCategory.OST_CableTray);
            //Category pipeFitting = Category.GetCategory(_doc, BuiltInCategory.OST_PipeFitting);
            if (element.Category.Id == pipe.Id)
            {
                return true;
            }
            else if (element.Category.Id == duct.Id)
            {
                return true;
            }
            else if (element.Category.Id == conduit.Id)
            {
                return true;
            }
            else if (element.Category.Id == tray.Id)
            {
                return true;
            }
            //else if (element.Category.Id == pipeFitting.Id)
            //{
            //    return true;
            //}
            return false;
        }

        public bool AllowReference(Reference refer, XYZ point)
        {
            return false;
        }
    }
    public class linkedPipeSelectionFilter : ISelectionFilter
    {
        private Document _doc;
        public linkedPipeSelectionFilter(Document doc)
        {
            this._doc = doc;
        }
        public bool AllowElement(Element element)
        {
            return true;
        }
        public bool AllowReference(Reference refer, XYZ point)
        {
            Category pipe = Category.GetCategory(_doc, BuiltInCategory.OST_PipeCurves);
            Category duct = Category.GetCategory(_doc, BuiltInCategory.OST_DuctCurves);
            Category conduit = Category.GetCategory(_doc, BuiltInCategory.OST_Conduit);
            Category tray = Category.GetCategory(_doc, BuiltInCategory.OST_CableTray);
            var elem = this._doc.GetElement(refer);
            if (elem != null && elem is RevitLinkInstance link)
            {
                var linkElem = link.GetLinkDocument().GetElement(refer.LinkedElementId);
                if (linkElem.Category.Id == pipe.Id)
                {
                    return true;
                }
                else if (linkElem.Category.Id == duct.Id)
                {
                    return true;
                }
                else if (linkElem.Category.Id == conduit.Id)
                {
                    return true;
                }
                else if (linkElem.Category.Id == tray.Id)
                {
                    return true;
                }
            }
            else
            {
                if (elem.Category.Id == pipe.Id)
                {
                    return true;
                }
                else if (elem.Category.Id == duct.Id)
                {
                    return true;
                }
                else if (elem.Category.Id == conduit.Id)
                {
                    return true;
                }
                else if (elem.Category.Id == tray.Id)
                {
                    return true;
                }
            }
            return false;
        }
    }
}

using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Autodesk.Revit.UI.Selection;
using System.Windows.Media.Imaging;
using System.IO;
using System.Drawing.Imaging;
using System.Drawing;


namespace PipeTagger.Common
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Picker
    {
        //public class All_Pipe_Selection_Filter : ISelectionFilter
        //{
        //    Document doc = null;
        //    public All_Pipe_Selection_Filter(Document document)
        //    {
        //        doc = document;
        //    }
        //    public bool AllowElement(Element elem)
        //    {
        //        if (elem.Category.Name == "管" || elem.Category.Name == "風管" || elem.Category.Name == "電管" || elem.Category.Name == "電纜架")
        //        {
        //            return true;
        //        }
        //        return true;
        //    }

        //    public bool AllowReference(Reference refe, XYZ position)
        //    {
        //        Element elem = doc.GetElement(refe);

        //        if (elem.GetType().Name == "RevitLinkInstance")
        //        {
        //            RevitLinkInstance revitlinkinstance = doc.GetElement(refe) as RevitLinkInstance;
        //            Document docLink = revitlinkinstance.GetLinkDocument();
        //            Element element = docLink.GetElement(refe.LinkedElementId);
        //            if (element.Category.Name == "管" || element.Category.Name == "風管" || element.Category.Name == "電管" || element.Category.Name == "電纜架")
        //            {
        //                return true;
        //            }
        //        }

        //        return false;
        //    }
        //}

        public class PipeSelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element element)
            {
                if (element.Category.Name == "管" || element.Category.Name == "風管" || element.Category.Name == "電管" || element.Category.Name == "電纜架")
                {
                    return true;
                }
                return false;
            }
            public bool AllowReference(Reference refer, XYZ point)
            {
                return false;
            }
        }

        public class LinkPipeSelectionFilter : ISelectionFilter
        {
            Document doc = null;
            public LinkPipeSelectionFilter(Document document)
            {
                doc = document;
            }
            public bool AllowElement(Element element)
            {
                return true;
            }
            public bool AllowReference(Reference refer, XYZ point)
            {

                RevitLinkInstance revitlinkinstance = doc.GetElement(refer) as RevitLinkInstance;
                Document docLink = revitlinkinstance.GetLinkDocument();
                Element element = docLink.GetElement(refer.LinkedElementId);
                if (element.Category.Name == "管" || element.Category.Name == "風管" || element.Category.Name == "電管" || element.Category.Name == "電纜架")
                {
                    return true;
                }
                return false;
            }
        }

        public static (bool isUp, bool isHorizontal, bool isLeft, IList<Reference> sorted_list) is_UD_HV_LR(UIDocument uidoc, List<Reference> pipes, XYZ pickPoint1, XYZ pickPoint2)
        {
            bool isUp = true;
            bool isHorizontal = true;
            bool isLeft = true;

            Document doc = uidoc.Document;

            #region 2.1.1 Get x, y location of pipes' project point by Pickpoint 1

            List<Double> xi = new List<Double>();
            List<Double> yi = new List<Double>();

            foreach (Reference refe in pipes)
            {
                Element elem = doc.GetElement(refe);

                //1. Get pipe element 
                // Case 1.1 Host element 
                if (elem.GetType().Name != "RevitLinkInstance")
                {
                    elem = doc.GetElement(refe);

                    LocationCurve locCurve = elem.Location as LocationCurve;
                    Line line = locCurve.Curve as Line;

                    xi.Add(line.Project(pickPoint1).XYZPoint.X);
                    yi.Add(line.Project(pickPoint1).XYZPoint.Y);

                    //MessageBox.Show((line.Project(pickPoint1).XYZPoint.X).ToString(), (line.Project(pickPoint1).XYZPoint.Y).ToString());
                }


                // Case 1.2 Linked instance
                // 此時選擇的外參reference仍是外參檔的reference，須轉型成外參檔裡的管
                else
                {
                    RevitLinkInstance linkInstance = elem as RevitLinkInstance; // 將選取的物件轉型為外參檔

                    //XYZ trans = linkInstance.GetTotalTransform().Origin;
                    Transform trans = linkInstance.GetTotalTransform();

                    Document docLink = linkInstance.GetLinkDocument(); // 指名新的document存放該外參檔資料
                    elem = docLink.GetElement(refe.LinkedElementId); // 從外參的document裡尋找欲選擇的管，轉型成element

                    LocationCurve locCurve = elem.Location as LocationCurve;

                    Curve line = locCurve.Curve as Curve;

                    Curve trans_line = line.CreateTransformed(trans);

                    xi.Add(trans_line.Project(pickPoint1).XYZPoint.X);
                    yi.Add(trans_line.Project(pickPoint1).XYZPoint.Y);

                    //MessageBox.Show((trans_line.Project(pickPoint1).XYZPoint.X).ToString(), 
                    //                (trans_line.Project(pickPoint1).XYZPoint.Y).ToString());

                }
            }

            //if (xi.Count != yi.Count) { MessageBox.Show("Picked points'x, y coord are not the same, something may happen"); }

            #endregion

            #region 2.1.1.A. Slope of pipes' midpoints-> tag text: isHorizontal ?

            // Case 1. no difference in x_i => pipes stack vertically => Horizontal text
            if (xi.Max() - xi.Min() < 0.001) { isHorizontal = true; }

            // Case 2. no difference in y_i => pipes stack horizontally => Vertical text
            else if (yi.Max() - yi.Min() < 0.001) { isHorizontal = false; }

            // Case 3. slope of midpoints of first two pipes (slope of every two pipes should be the same) -> Absolute value 
            else
            {
                List<double> slopes = new List<double>();
                foreach (int coor1 in Enumerable.Range(0, xi.Count))
                {
                    foreach (int coor2 in Enumerable.Range(0, xi.Count))
                    {
                        if (coor2 != coor1)
                        {
                            slopes.Add(Math.Abs((yi[coor1] - yi[coor2]) / (xi[coor1] - xi[coor2])));
                        }
                    }
                }

                if ((slopes.Max() - slopes.Min()) > 0.001)
                {
                    //MessageBox.Show("生成的標籤可能無法對齊顯示,請注意第一個點擊點.");
                    //MessageBox.Show("The first picked point cannot project to all pipes, the tages may not be aligned!!!");
                    //MessageBox.Show($"Max slope: {slopes.Max()}, Min slope: {slopes.Min()}");
                }


                double slope = slopes.Average();
                if (slope >= 1) { isHorizontal = true; }
                else { isHorizontal = false; }
            }
            #endregion

            #region 2.1.1.B. pickPoint1 -> isUp ?

            if (isHorizontal) { isUp = (pickPoint1.Y > yi.Average()) ? true : false; }  //Case Horizontal text:  1. Pick.Y > Y_avg => UP  2. Pick.Y < Y_avg => Down 
            else { isUp = (pickPoint1.X > xi.Average()) ? true : false; }                //Case vertical text:    1. Pick.X > X_avg => UP  2. Pick.X < X_avg => Down   

            #endregion

            #region 2.1.1.C pickPoint2 -> isLeft?

            if (isHorizontal) { isLeft = (pickPoint1.X > pickPoint2.X) ? true : false; } //Case Horizontal text:  1. Pick1.X > Pick2.X => Left  2. Pick1.X < Pick2.X => Right
            else { isLeft = (pickPoint1.Y > pickPoint2.Y) ? true : false; }             //Case vertical text:    1. Pick1.Y > Pick2.Y => Left  2. Pick1.Y > Pick2.Y => Right
            #endregion

            #region 2.1.2 Sorted Tag's order

            // 4 Cases in total: 
            // using y coor. : U-H / D-H 
            // using x coor. : U-V / D-V 

            pipes.Sort((ref1, ref2) =>
            {



                Element elem1 = Common_Tool.host_link_element_All_in_one(ref1, doc, "Element");
                Element elem2 = Common_Tool.host_link_element_All_in_one(ref2, doc, "Element");

                #region Replace by common tool.Get_host_or_link_element


                //Element elem1 = doc.GetElement(ref1);
                //Element elem2 = doc.GetElement(ref2);

                //if (elem1.GetType().Name != "RevitLinkInstance"){
                //    elem1 = doc.GetElement(ref1);}
                //else{
                //    RevitLinkInstance linkInstance = elem1 as RevitLinkInstance; // 將選取的物件轉型為外參檔      
                //    Document docLink = linkInstance.GetLinkDocument(); // 指名新的document存放該外參檔資料
                //    elem1 = docLink.GetElement(ref1.LinkedElementId); // 從外參的document裡尋找欲選擇的管，轉型成element
                //}

                //if (elem2.GetType().Name != "RevitLinkInstance")
                //{
                //    elem2 = doc.GetElement(ref2);
                //}
                //else
                //{
                //    RevitLinkInstance linkInstance = elem2 as RevitLinkInstance; // 將選取的物件轉型為外參檔      
                //    Document docLink = linkInstance.GetLinkDocument(); // 指名新的document存放該外參檔資料
                //    elem2 = docLink.GetElement(ref2.LinkedElementId); // 從外參的document裡尋找欲選擇的管，轉型成element
                //}
                #endregion

                Line line1 = (elem1.Location as LocationCurve).Curve as Line;
                Line line2 = (elem2.Location as LocationCurve).Curve as Line;

                //XYZ Porj_pt1 = line1.Project(pickPoint1).XYZPoint;
                //XYZ Porj_pt2 = line2.Project(pickPoint1).XYZPoint;

                XYZ Porj_pt1 = Common_Tool.host_link_element_All_in_one(ref1, doc, "Projection", pickPoint1);
                XYZ Porj_pt2 = Common_Tool.host_link_element_All_in_one(ref2, doc, "Projection", pickPoint1);

                int comparer = 0;

                if (isUp & isHorizontal) { comparer = (Porj_pt1.Y).CompareTo(Porj_pt2.Y); }
                if (!isUp & isHorizontal) { comparer = (Porj_pt2.Y).CompareTo(Porj_pt1.Y); }

                if (isUp & !isHorizontal) { comparer = (Porj_pt1.X).CompareTo(Porj_pt2.X); }
                if (!isUp & !isHorizontal) { comparer = (Porj_pt2.X).CompareTo(Porj_pt1.X); }


                //MessageBox.Show(comparer.ToString()); => -1, 1 , 0

                return comparer;


            });



            #endregion

            return (isUp, isHorizontal, isLeft, pipes);
        }

    }
    public class Placer
    {
#if RELEASE2019
         public static DisplayUnitType unitType = DisplayUnitType.DUT_MILLIMETERS;
#else
        public static ForgeTypeId unitType = UnitTypeId.Millimeters;
#endif


        public static Result CreateSomething(UIDocument uidoc, Document doc, ref string message, string place_type)
        {
            try
            {

                List<Reference> pipes = new List<Reference>();
                using (TransactionGroup transactionGroup = new TransactionGroup(doc))
                {
                    transactionGroup.Start("Select & Create preview line");
                    using (Transaction t = new Transaction(doc, "Select & Create preview line"))
                    {
                        t.Start();

                        #region 1. Pick pipe and put into a list "pipes"

                        //Filter for internal and linked "Pipe"
                        ISelectionFilter selFilter = new Picker.PipeSelectionFilter();
                        ISelectionFilter selFilter_link = new Picker.LinkPipeSelectionFilter(doc);

                        #region [To do] Pick host and link at the same time

                        //ISelectionFilter selFilter_All = new Picker.All_Pipe_Selection_Filter(doc);
                        //IList<Reference> refElemLinked_all = uidoc.Selection.PickObjects(ObjectType.PointOnElement, selFilter_All, "3.請點選要連續標註的管，結束選取按完成");
                        //IList<Reference> refElemLinked_all = uidoc.Selection.PickObjects(ObjectType.Element, selFilter, "1.請點選要連續標註的管(專案內)，結束選取按完成", uidoc.Selection.PickObjects(ObjectType.LinkedElement, selFilter_link, "2.請點選要連續標註的管(外參)，結束選取按完成"));

                        //MessageBox.Show(refElemLinked_all.Count.ToString());
                        //IList<Element> refElem_PickBox = uidoc.Selection.PickElementsByRectangle();

                        //string outtt = "";
                        //foreach (Element elem in refElem_PickBox)
                        //{
                        //    outtt += elem.Name + "\n";
                        //}

                        //MessageBox.Show(outtt);
                        //MessageBox.Show(refElem_PickBox.Count.ToString());

                        // 依序轉型成Element，存入pipes
                        //foreach (Element element in refElem_PickBox) //Concat  = refElem + refElemLinked
                        //{

                        //    pipes.Add(refe);
                        //}
                        #endregion

                        //Allow user to pick pipe only
                        IList<Reference> refElem = uidoc.Selection.PickObjects(ObjectType.Element, selFilter, "1.請點選要連續標註的管(專案內)，結束選取按完成");


                        IList<Reference> refElemLinked = null;

                        if (place_type != "Spot")
                        {
                            refElemLinked = uidoc.Selection.PickObjects(ObjectType.LinkedElement, selFilter_link, "2.請點選要連續標註的管(外參)，結束選取按完成");
                        }

                        //Pick point to place tag
                        XYZ pickPoint1 = uidoc.Selection.PickPoint("3.請點選標籤起始位置");
                        XYZ temp_proj_on_pipe = new XYZ();
                        if (refElem.Count() != 0)
                        {
                            temp_proj_on_pipe = Common_Tool.host_link_element_All_in_one(refElem[0], doc, "Projection", pickPoint1);
                        }
                        else
                        {
                            temp_proj_on_pipe = Common_Tool.host_link_element_All_in_one(refElemLinked[0], doc, "Projection", pickPoint1);
                        }
                        // Create preview line 


                        XYZ P1 = new XYZ(pickPoint1.X, pickPoint1.Y, 0);
                        XYZ P2 = new XYZ(temp_proj_on_pipe.X, temp_proj_on_pipe.Y, 0);
                        Line L1 = Line.CreateBound(P1, P2);
                        var temp_line = doc.Create.NewDetailCurve(doc.ActiveView, L1);



                        XYZ pickPoint2 = uidoc.Selection.PickPoint("4.請選擇標籤向左長或向右長(往左或右任意點選)");

                        doc.Delete(temp_line.Id);

                        // 依序轉型成Element，存入pipes

                        if (refElem != null)
                        {
                            foreach (Reference refe in refElem)
                            {
                                pipes.Add(refe);
                            }
                        }

                        if (place_type != "Spot")
                        {
                            foreach (Reference refe in refElemLinked)
                            {
                                pipes.Add(refe);
                            }
                        }


                        #endregion

                        (bool isUp, bool isHorizontal, bool isLeft, IList<Reference> sorted_list) = Picker.is_UD_HV_LR(uidoc, pipes, pickPoint1, pickPoint2);

                        #region [DEBUG only] Check UP/Left/Hoeizontal
                        //string UD = isUp ? "U" : "D";
                        //string LR = isLeft ? "L" : "R";
                        //string HV = isHorizontal ? "H" : "V";
                        //MessageBox.Show($"{UD}{LR}{HV}");
                        #endregion

                        // Initial tag location for recursive location

                        XYZ first_proj_on_pipe = Common_Tool.host_link_element_All_in_one(sorted_list[0], doc, "Projection", pickPoint1);
                        Line normLine = Line.CreateBound(first_proj_on_pipe, pickPoint1);
                        XYZ lineDir = normLine.Direction;
                        XYZ normal_dir = null;
                        normal_dir = new XYZ(lineDir.X, lineDir.Y, 0);
                        normal_dir = normal_dir.Normalize();

                        //normal_dir = new XYZ(normal_dir.X, normal_dir.Y, 0);

                        int num_placement = 0;
                        double adjust = UnitUtils.ConvertToInternalUnits(Properties.Settings.Default.label_dist_mm, unitType);



                        //For pip in pipes -> Get location with Host element or linked instance

                        t.Commit();

                        foreach (Reference refe in sorted_list)
                        {

                            //1. Get Element from host / linked
                            Element elem = Common_Tool.host_link_element_All_in_one(refe, doc, "Element");

                            //2. Get pipe's location
                            XYZ proj_on_pipe = Common_Tool.host_link_element_All_in_one(refe, doc, "Projection", pickPoint1);

                            //XYZ extented_pt = new XYZ(num_placement * Properties.Settings.Default.label_dist_mm / 304.8 * normal_dir.X,
                            //                          num_placement * Properties.Settings.Default.label_dist_mm / 304.8 * normal_dir.Y,
                            //                          0.0);
                            XYZ extented_pt = new XYZ(num_placement * adjust * normal_dir.X, num_placement * adjust * normal_dir.Y, 0.0);

                            //3. Place Independent Tag
                            Placer.Create_Tag_or_Spot(doc, elem, proj_on_pipe, extented_pt + pickPoint1, refe, isUp, isHorizontal, isLeft, place_type, pickPoint1);
                            num_placement += 1;
                        }
                    }
                    transactionGroup.Assimilate();
                }

                #region Test
                //string pipes_xyzs = "";
                //int count = 1;
                //foreach (XYZ xyz in xyzs)
                //{
                //    pipes_xyzs += "####\n" + (count / 2).ToString() + "\n";
                //    pipes_xyzs += xyz.ToString() + "\n";
                //    count += 1;
                //}
                //MessageBox.Show("pipes.Count" + pipes.Count.ToString());
                //MessageBox.Show("xyzss.Count" + xyzs.Count.ToString());
                //MessageBox.Show(pipes_xyzs);
                //string output = "";
                //foreach (Reference elem in refElem) { output += elem.ElementId.ToString() + "\n"; }
                //foreach (Reference elem in refElemLinked) { output += elem.ElementId.ToString() + "\n"; }
                //MessageBox.Show(output);

                #endregion

                return Result.Succeeded;
            }

            catch (Exception e)
            {
                //MessageBox.Show(e.Message.ToString());
                //message = e.Message;
                return Result.Failed;
            }
        }

        public static IndependentTag Create_Tag_or_Spot(Document document, Element elem, XYZ on_pipe, XYZ tag_start, Reference refe, bool isUp, bool isHorizontal, bool isLeft, string place_type, XYZ pickpoint1)
        {
            using (Transaction t = new Transaction(document, "Select & Create preview line"))
            {
                t.Start("建立標籤");
                // Make sure active view is not a 3D view
                var view = document.ActiveView;

                // Define tag mode and tag orientation for new tag
                TagMode tagMode = TagMode.TM_ADDBY_CATEGORY;

                // Add the tag to the projected point from Click_Pt to picked pipe of the wall
                Reference elemRef = Common_Tool.host_link_element_All_in_one(refe, document, "Reference");
                IndependentTag newTag;

                #region Create tag
                newTag = IndependentTag.Create(document, view.Id, elemRef, true, tagMode, TagOrientation.Horizontal, on_pipe);

                if (null == newTag) throw new Exception("Create IndependentTag Failed.");

                newTag.ChangeTypeId(getPipeTagId(document, refe));

                double Tag_text_length;

                // All diff are measure in foot
                // 1 ft = 12 * 25.4 = 304.8 mm


                newTag.LeaderEndCondition = LeaderEndCondition.Free;
                if (pickpoint1.X < on_pipe.X)
                {
#if RELEASE2023
                    newTag.SetLeaderElbow(refe, new XYZ(tag_start.X + 2 * (on_pipe.X - tag_start.X), tag_start.Y + 2 * (on_pipe.Y - tag_start.Y), tag_start.Z));
#else
;                    newTag.LeaderElbow = new XYZ(tag_start.X + 2 * (on_pipe.X - tag_start.X), tag_start.Y + 2 * (on_pipe.Y - tag_start.Y), tag_start.Z);
#endif
                }
                else
                {
#if RELEASE2023
                    newTag.SetLeaderElbow(refe, tag_start);
#else
                    newTag.LeaderElbow = tag_start;
#endif
#if RELEASE2023
#else
#endif
                }
#if RELEASE2023
                newTag.SetLeaderEnd(refe, on_pipe);
#else
                newTag.LeaderEnd = on_pipe;
#endif

                XYZ header_pos = null;
#if RELEASE2023
                header_pos = newTag.GetLeaderElbow(refe);
#else
                header_pos = newTag.LeaderElbow;
#endif
                //double label_length = Properties.Settings.Default.label_length_mm / 304.8;
                double label_length = UnitUtils.ConvertToInternalUnits(Properties.Settings.Default.label_length_mm, unitType);

                if (isHorizontal) { header_pos += new XYZ(label_length, 0.0, 0.0); }
                if (!isHorizontal) { header_pos += new XYZ(0.0, label_length, 0.0); }

                newTag.TagHeadPosition = header_pos;

                t.Commit();

                t.Start("調整標籤長度");

                // Default orientation = Horizontal

                Tag_text_length = newTag.get_BoundingBox(view).Max.X - newTag.TagHeadPosition.X;
                if (pickpoint1.X < on_pipe.X)
                {
#if RELEASE2023
                    newTag.SetLeaderElbow(refe, tag_start);
#else
                    newTag.LeaderElbow = tag_start;
#endif
                }

                if (place_type != "Spot")
                {
                    if (isLeft & isHorizontal)
                    {
                        header_pos = tag_start + new XYZ(-Tag_text_length - label_length, 0.0, 0.0);
                        newTag.TagHeadPosition = header_pos;
                        //Tag_text_length = newTag.TagHeadPosition.X - newTag.get_BoundingBox(view).Min.X;
                        //header_pos += new XYZ(-3 * label_length - Tag_text_length, 0.0, 0.0);
                        //newTag.TagHeadPosition = header_pos;
                        //header_pos += new XYZ(-2 * label_length - Tag_text_length, 0.0, 0.0);
                        //newTag.TagHeadPosition = header_pos;
                    }

                    if (isLeft & !isHorizontal)
                    {
                        header_pos = tag_start + new XYZ(0.0, -Tag_text_length - label_length, 0.0);
                        newTag.TagHeadPosition = header_pos;
                        //Tag_text_length = newTag.get_BoundingBox(view).Max.Y - newTag.get_BoundingBox(view).Min.Y;
                        //MessageBox.Show(Tag_text_length.ToString());
                        //header_pos += new XYZ(-label_length, -label_length - Tag_text_length, 0.0);
                    }
                    if (!isLeft & !isHorizontal)
                    {
                        header_pos = tag_start + new XYZ(0.0, label_length, 0.0);
                        newTag.TagHeadPosition = header_pos;

                    }
                }

                if (!isHorizontal)
                {
                    newTag.TagOrientation = TagOrientation.Vertical;
#if RELEASE2023
                    newTag.TagHeadPosition = new XYZ(newTag.GetLeaderElbow(refe).X, newTag.TagHeadPosition.Y, newTag.TagHeadPosition.Z);
#else
       newTag.TagHeadPosition = new XYZ(newTag.LeaderElbow.X, newTag.TagHeadPosition.Y, newTag.TagHeadPosition.Z);
#endif
                }
                else
                {
#if RELEASE2023

                    newTag.TagHeadPosition = new XYZ(newTag.TagHeadPosition.X, newTag.GetLeaderElbow(refe).Y, newTag.TagHeadPosition.Z);
#else
                    newTag.TagHeadPosition = new XYZ(newTag.TagHeadPosition.X, newTag.LeaderElbow.Y, newTag.TagHeadPosition.Z);
#endif

                }


                newTag.LeaderEndCondition = LeaderEndCondition.Attached;



                if (place_type == "Spot")
                {
                    SpotDimension new_spot = document.Create.NewSpotElevation(view, elemRef, on_pipe, tag_start, header_pos, on_pipe, true);

                    document.Delete(newTag.Id);
                }

                t.Commit();

                #endregion

                //Tag_text_length = newTag.get_BoundingBox(view).Max.X - newTag.TagHeadPosition.X;

                return newTag;
            }
        }

        public static ElementId getPipeTagId(Document doc, Reference refe)
        {
            //Element pipe = doc.GetElement(refe);

            Element pipe = Common_Tool.host_link_element_All_in_one(refe, doc, "Element");
            ElementId pipeTagId = null;

            // Create a list to store the tags
            IList<Element> pipeTags = new FilteredElementCollector(doc).OfClass(typeof(Family)).Where(x =>
                                                                                                            x.Name == Properties.Settings.Default.PI_1st_Selected.ToString() ||
                                                                                                            x.Name == Properties.Settings.Default.DT_C_1st_Selected.ToString() ||
                                                                                                            x.Name == Properties.Settings.Default.DT_R_1st_Selected.ToString() ||
                                                                                                            x.Name == Properties.Settings.Default.CN_1st_Selected.ToString() ||
                                                                                                            x.Name == Properties.Settings.Default.CT_1st_Selected.ToString()).ToList();

            // Match the [ Pipe.Category.Name ] & [ Pipe Tag ]   
            foreach (Element elem in pipeTags)
            {
                Family family = elem as Family;
                ISet<ElementId> familySymbolIds = family.GetFamilySymbolIds();



                if (pipe.Category.Name is "管")
                {
                    if (family.Name == Properties.Settings.Default.PI_1st_Selected.ToString())  // e.g. "L_圓形管標籤_標註符號"
                    {
                        foreach (ElementId elemId in familySymbolIds)
                        {
                            FamilySymbol familySymbol = doc.GetElement(elemId) as FamilySymbol;


                            if (familySymbol.Name == Properties.Settings.Default.PI_2nd_Selected) // e.g. "標稱+系統+BOP/COP"
                            {
                                pipeTagId = familySymbol.Id;
                                break;
                            }
                        }
                        break;
                    }
                }
                else if (pipe.Category.Name is "風管")
                {
                    Autodesk.Revit.DB.Mechanical.Duct duct = pipe as Autodesk.Revit.DB.Mechanical.Duct;
                    Autodesk.Revit.DB.Mechanical.DuctType ductType = duct.DuctType;
                    if (ductType.FamilyName is "圓形風管")
                    {
                        if (family.Name == Properties.Settings.Default.DT_C_1st_Selected.ToString())
                        {
                            foreach (ElementId elemId in familySymbolIds)
                            {
                                FamilySymbol familySymbol = doc.GetElement(elemId) as FamilySymbol;
                                if (familySymbol.Name == Properties.Settings.Default.DT_C_2nd_Selected)
                                {
                                    pipeTagId = familySymbol.Id;
                                    break;
                                }
                            }
                            break;
                        }
                    }
                    else
                    {
                        if (family.Name == Properties.Settings.Default.DT_R_1st_Selected.ToString())
                        {
                            foreach (ElementId elemId in familySymbolIds)
                            {
                                FamilySymbol familySymbol = doc.GetElement(elemId) as FamilySymbol;
                                if (familySymbol.Name == Properties.Settings.Default.DT_R_2nd_Selected) //is "WxH+系統+BOD")
                                {
                                    pipeTagId = familySymbol.Id;
                                    break;
                                }
                            }
                            break;
                        }
                    }
                }
                else if (pipe.Category.Name is "電管")
                {
                    if (family.Name == Properties.Settings.Default.CN_1st_Selected.ToString())
                    {
                        foreach (ElementId elemId in familySymbolIds)
                        {
                            FamilySymbol familySymbol = doc.GetElement(elemId) as FamilySymbol;
                            if (familySymbol.Name == Properties.Settings.Default.CN_2nd_Selected) //is "材質+標稱+BOP")
                            {
                                pipeTagId = familySymbol.Id;
                                break;
                            }
                        }
                        break;
                    }
                }
                else if (pipe.Category.Name is "電纜架")
                {
                    if (family.Name == Properties.Settings.Default.CT_1st_Selected.ToString())
                    {
                        foreach (ElementId elemId in familySymbolIds)
                        {
                            FamilySymbol familySymbol = doc.GetElement(elemId) as FamilySymbol;
                            if (familySymbol.Name == Properties.Settings.Default.CT_2nd_Selected) // is "服務類型+標稱+BOD")
                            {
                                pipeTagId = familySymbol.Id;
                                break;
                            }
                        }
                        break;
                    }
                }
                else
                    MessageBox.Show("此物件不是任一種管");
            }
            return pipeTagId;
        }

    }
    public class Common_Tool
    {
        public static dynamic host_link_element_All_in_one(Reference refe, Document doc, string return_type, XYZ pickPoint1 = null)
        {

            ///<example>
            ///
            /// return [Element]    ==>  Common_Tool.host_link_element_All_in_one(refe, document, "Element");
            /// return [Reference]  ==>  Common_Tool.host_link_element_All_in_one(refe, document, "Reference");
            /// return [XYZ]        ==>  Common_Tool.host_link_element_All_in_one(ref1, doc, "Projection", pickPoint1);
            /// 
            /// </example>

            string[] return_types = { "Reference", "Element", "Projection" };

            if (return_types.Contains(return_type) == false)
            {
                MessageBox.Show("From func [host_link_element_All_in_one], please input return_type either 'Reference', 'Element', or 'Projection'");
                return 0;
            }

            Element elem = doc.GetElement(refe);

            LocationCurve locCurve;
            Line line;

            if (elem.GetType().Name != "RevitLinkInstance")
            {
                elem = doc.GetElement(refe);
                double slope = 0;
                //水管
                if ((BuiltInCategory)elem.Category.Id.IntegerValue == BuiltInCategory.OST_PipeCurves)
                {
                    slope = elem.get_Parameter(BuiltInParameter.RBS_PIPE_SLOPE).AsDouble();
                }
                //電管
                if ((BuiltInCategory)elem.Category.Id.IntegerValue == BuiltInCategory.OST_Conduit)
                {
                    slope = elem.get_Parameter(BuiltInParameter.RBS_CURVE_SLOPE).AsDouble();
                }
                //風管
                if ((BuiltInCategory)elem.Category.Id.IntegerValue == BuiltInCategory.OST_DuctCurves)
                {
                    slope = elem.get_Parameter(BuiltInParameter.RBS_DUCT_SLOPE).AsDouble();
                }
                //電纜架
                if ((BuiltInCategory)elem.Category.Id.IntegerValue == BuiltInCategory.OST_CableTray)
                {
                    slope = elem.get_Parameter(BuiltInParameter.RBS_CURVE_SLOPE).AsDouble();
                }
                locCurve = elem.Location as LocationCurve;
                line = locCurve.Curve as Line;
                if (slope != 0)
                {
                    XYZ point = new XYZ(line.Tessellate()[1].X, line.Tessellate()[1].Y, line.Tessellate()[0].Z);
                    line = Line.CreateBound(line.Tessellate()[0], point);
                }
                if (return_type == "Reference")
                {
                    return new Reference(elem);
                }
            }
            else
            {
                RevitLinkInstance linkInstance = elem as RevitLinkInstance; // 將選取的物件轉型為外參檔
                Document docLink = linkInstance.GetLinkDocument(); // 指名新的document存放該外參檔資料
                elem = docLink.GetElement(refe.LinkedElementId); // 從外參的document裡尋找欲選擇的管，轉型成element

                Transform Link_trans = linkInstance.GetTotalTransform();
                locCurve = elem.Location as LocationCurve;
                line = locCurve.Curve as Line;
                double slope = 0;
                //水管
                if ((BuiltInCategory)elem.Category.Id.IntegerValue == BuiltInCategory.OST_PipeCurves)
                {
                    slope = elem.get_Parameter(BuiltInParameter.RBS_PIPE_SLOPE).AsDouble();
                }
                //電管
                if ((BuiltInCategory)elem.Category.Id.IntegerValue == BuiltInCategory.OST_Conduit)
                {
                    slope = elem.get_Parameter(BuiltInParameter.RBS_CURVE_SLOPE).AsDouble();
                }
                //風管
                if ((BuiltInCategory)elem.Category.Id.IntegerValue == BuiltInCategory.OST_DuctCurves)
                {
                    slope = elem.get_Parameter(BuiltInParameter.RBS_DUCT_SLOPE).AsDouble();
                }
                //電纜架
                if ((BuiltInCategory)elem.Category.Id.IntegerValue == BuiltInCategory.OST_CableTray)
                {
                    slope = elem.get_Parameter(BuiltInParameter.RBS_CURVE_SLOPE).AsDouble();
                }
                if (slope != 0)
                {
                    XYZ point = new XYZ(line.Tessellate()[1].X, line.Tessellate()[1].Y, line.Tessellate()[0].Z);
                    line = Line.CreateBound(line.Tessellate()[0], point);
                }
                line = (Line)line.CreateTransformed(Link_trans);

                if (return_type == "Reference")
                {
                    return new Reference(docLink.GetElement(refe.LinkedElementId)).CreateLinkReference(linkInstance);
                }

            }

            if (return_type == "Element") { return elem; }

            if (return_type == "Projection")
            {
                if (pickPoint1 == null) { MessageBox.Show("From func [host_link_element_All_in_one], please input the pickpoint if you want to get the project point"); }
                try { return line.Project(pickPoint1).XYZPoint; }
                catch (Exception e) { MessageBox.Show(e.Message.ToString()); }
            }

            return "haha";

        }

        #region 3 func combined as host_link_element_All_in_one

        //public static Element Get_host_or_link_element(Reference refe, Document doc)
        //{
        //    Element elem = doc.GetElement(refe);

        //    if (elem.GetType().Name != "RevitLinkInstance")
        //    {
        //        elem = doc.GetElement(refe);
        //    }
        //    else
        //    {
        //        RevitLinkInstance linkInstance = elem as RevitLinkInstance; // 將選取的物件轉型為外參檔      
        //        Document docLink = linkInstance.GetLinkDocument(); // 指名新的document存放該外參檔資料
        //        elem = docLink.GetElement(refe.LinkedElementId); // 從外參的document裡尋找欲選擇的管，轉型成element
        //        //MessageBox.Show(elem.Category.Name);
        //    }

        //    return elem;
        //}

        //public static Reference Create_new_reference(Reference refe, Document doc)
        //{
        //    Element elem = doc.GetElement(refe);

        //    // Case 1: Host element
        //    if (elem.GetType().Name != "RevitLinkInstance")
        //    {
        //        return new Reference(elem);
        //    }
        //    // Case 2: Linked element
        //    else
        //    {

        //        RevitLinkInstance linkInstance = doc.GetElement(refe) as RevitLinkInstance;
        //        Document docLink = linkInstance.GetLinkDocument();

        //        Reference elemRef = new Reference(docLink.GetElement(refe.LinkedElementId)).CreateLinkReference(linkInstance);

        //        return elemRef;

        //    }

        //}

        //public static XYZ Get_transformed_project_pt(Reference refe, Document doc, XYZ pickPoint1)
        //{
        //    Element elem = doc.GetElement(refe);

        //    XYZ proj_on_pipe;
        //    LocationCurve locCurve;
        //    Line line;

        //    if (elem.GetType().Name != "RevitLinkInstance")
        //    {
        //        elem = doc.GetElement(refe);
        //        locCurve = elem.Location as LocationCurve;
        //        line = locCurve.Curve as Line;
        //        proj_on_pipe = line.Project(pickPoint1).XYZPoint;
        //    }
        //    else
        //    {
        //        RevitLinkInstance linkInstance = elem as RevitLinkInstance; // 將選取的物件轉型為外參檔
        //        XYZ trans = linkInstance.GetTotalTransform().Origin;
        //        Transform Link_trans = linkInstance.GetTotalTransform();
        //        Document docLink = linkInstance.GetLinkDocument(); // 指名新的document存放該外參檔資料
        //        elem = docLink.GetElement(refe.LinkedElementId); // 從外參的document裡尋找欲選擇的管，轉型成element
        //        locCurve = elem.Location as LocationCurve;
        //        line = locCurve.Curve as Line;
        //        Curve trans_line = line.CreateTransformed(Link_trans);

        //        proj_on_pipe = trans_line.Project(pickPoint1).XYZPoint;
        //    }

        //    return proj_on_pipe;

        //}


        #endregion

    }

    public class UI_Tool
    {
        public static List<String> Get_Tag_2nd_options(String family_Name, Document doc)
        {
            List<string> TypeOptions = new List<string>();


            //IList<Element> pipeTags = new FilteredElementCollector(doc).OfClass(typeof(Family)).Where(x => x.Name is "L_圓形管標籤_標註符號" ||
            //                                                                                               x.Name is "L_圓形風管標籤_標註符號" ||
            //                                                                                               x.Name is "L_方形風管標籤_標註符號" ||
            //                                                                                               x.Name is "L_電管標籤_標註符號" ||
            //                                                                                               x.Name is "L_電纜架標籤_標註符號").ToList();

            IList<Element> pipeTags = new FilteredElementCollector(doc).OfClass(typeof(Family)).Where(x => x.Name.Contains(family_Name)).ToList();

            if (pipeTags.Count == 0) { return TypeOptions; }
            //MessageBox.Show("Get_Tag_2nd_options!!!");

            foreach (Element elem in pipeTags)
            {
                Family family = elem as Family;
                ISet<ElementId> familySymbolIds = family.GetFamilySymbolIds();

                if (family.Name == family_Name)
                {
                    foreach (ElementId elemId in familySymbolIds)
                    {
                        FamilySymbol familySymbol = doc.GetElement(elemId) as FamilySymbol;
                        TypeOptions.Add(familySymbol.Name);
                    }
                }

                #region Reaplce by above code
                //if (type is "管")
                //{
                //    if (family.Name is "L_圓形管標籤_標註符號")
                //    {
                //        foreach (ElementId elemId in familySymbolIds)
                //        {
                //            FamilySymbol familySymbol = doc.GetElement(elemId) as FamilySymbol;
                //            TypeOptions.Add(familySymbol.Name);
                //            //MessageBox.Show("L_圓形管標籤_標註符號" + familySymbol.Name);
                //        }
                //    }
                //}

                //else if (type is "圓風管")
                //{
                //    if (family.Name is "L_圓形風管標籤_標註符號")
                //    {
                //        foreach (ElementId elemId in familySymbolIds)
                //        {
                //            FamilySymbol familySymbol = doc.GetElement(elemId) as FamilySymbol;
                //            TypeOptions.Add(familySymbol.Name);
                //            //MessageBox.Show("L_圓形風管標籤_標註符號" + familySymbol.Name);
                //        }
                //    }
                //}
                //else if (type is "方風管")
                //{
                //    if (family.Name is "L_方形風管標籤_標註符號")
                //    {
                //        foreach (ElementId elemId in familySymbolIds)
                //        {
                //            FamilySymbol familySymbol = doc.GetElement(elemId) as FamilySymbol;
                //            TypeOptions.Add(familySymbol.Name);
                //            //MessageBox.Show("L_方形風管標籤_標註符號" + familySymbol.Name);
                //        }
                //    }
                //}
                //else if (type is "電管")
                //{
                //    if (family.Name is "L_電管標籤_標註符號")
                //    {
                //        foreach (ElementId elemId in familySymbolIds)
                //        {
                //            FamilySymbol familySymbol = doc.GetElement(elemId) as FamilySymbol;
                //            TypeOptions.Add(familySymbol.Name);
                //            //MessageBox.Show("L_電管標籤_標註符號" + familySymbol.Name);
                //        }
                //    }
                //}
                //else if (type is "電纜架")
                //{
                //    if (family.Name is "L_電纜架標籤_標註符號")
                //    {
                //        foreach (ElementId elemId in familySymbolIds)
                //        {
                //            FamilySymbol familySymbol = doc.GetElement(elemId) as FamilySymbol;
                //            TypeOptions.Add(familySymbol.Name);
                //            //MessageBox.Show("L_電纜架標籤_標註符號" + familySymbol.Name);
                //        }
                //    }
                //}

                #endregion

                //else MessageBox.Show("此物件不是任一種管");
            }
            return TypeOptions;
        }

        public static List<String> Get_Tag_1st_options(String type, Document doc)
        {
            List<string> TypeOptions = new List<string>();

            if (type is "管")
            {
                IList<Element> pipeTags = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeTags).OfClass(typeof(FamilySymbol)).WhereElementIsElementType().ToElements();
                List<String> option = new List<String>();

                foreach (Element elem in pipeTags)
                {
                    option.Add(((ElementType)elem).FamilyName.ToString());
                }
                option = option.Distinct().ToList();
                //foreach (String opt in option) { MessageBox.Show(opt); }
                return option;
            }

            else if (type is "圓風管")
            {
                IList<Element> pipeTags = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctTags).OfClass(typeof(FamilySymbol)).WhereElementIsElementType().ToElements();
                List<String> option = new List<String>();

                foreach (Element elem in pipeTags)
                {
                    option.Add(((ElementType)elem).FamilyName.ToString());
                }
                option = option.Distinct().ToList();
                //foreach (String opt in option) { MessageBox.Show(opt); }
                return option;
            }
            else if (type is "方風管")
            {
                IList<Element> pipeTags = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctTags).OfClass(typeof(FamilySymbol)).WhereElementIsElementType().ToElements();
                List<String> option = new List<String>();

                foreach (Element elem in pipeTags)
                {
                    option.Add(((ElementType)elem).FamilyName.ToString());
                }
                option = option.Distinct().ToList();
                //foreach (String opt in option) { MessageBox.Show(opt); }
                return option;
            }
            else if (type is "電管")
            {
                IList<Element> pipeTags = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ConduitTags).OfClass(typeof(FamilySymbol)).WhereElementIsElementType().ToElements();
                List<String> option = new List<String>();

                foreach (Element elem in pipeTags)
                {
                    option.Add(((ElementType)elem).FamilyName.ToString());
                }
                option = option.Distinct().ToList();
                //foreach (String opt in option) { MessageBox.Show(opt); }
                return option;
            }
            else if (type is "電纜架")
            {
                IList<Element> pipeTags = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_CableTrayTags).OfClass(typeof(FamilySymbol)).WhereElementIsElementType().ToElements();
                List<String> option = new List<String>();

                foreach (Element elem in pipeTags)
                {
                    option.Add(((ElementType)elem).FamilyName.ToString());
                }
                option = option.Distinct().ToList();
                //foreach (String opt in option) { MessageBox.Show(opt); }
                return option;
            }
            //else MessageBox.Show("此物件不是任一種管");
            return TypeOptions;
        }

        public static BitmapSource GetImageSource(Image img)
        {
            //製作一個function專門來處理圖片
            BitmapImage bmp = new BitmapImage();

            using (MemoryStream ms = new MemoryStream())
            {
                img.Save(ms, ImageFormat.Png);
                ms.Position = 0;

                bmp.BeginInit();

                bmp.CacheOption = BitmapCacheOption.OnLoad;
                bmp.UriSource = null;
                bmp.StreamSource = ms;

                bmp.EndInit();
            }

            return bmp;
        }

    }
}
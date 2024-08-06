#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.ExtensibleStorage;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Autodesk.Revit.DB.Structure;
using System.Windows.Forms;
using System.Threading;
#endregion

namespace CEC_WallCast
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class NumberRuleSet : IExternalCommand
    {
#if RELEASE2019
        public static DisplayUnitType unitType = DisplayUnitType.DUT_MILLIMETERS;
#else
        public ForgeTypeId unitType = UnitTypeId.Millimeters;
#endif
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Counter.count += 1;
            string messageOut = "";

            UIApplication uiapp = commandData.Application;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            Autodesk.Revit.DB.View activeView = doc.ActiveView;

            
            //外部儲存功能測試

            using (Transaction trans = new Transaction(doc))
            {
                trans.Start("UserData測試");
                //創建SchemaBuilder
                string dataToStore = "醜";
                SchemaBuilder schemaBuilder = new SchemaBuilder(new Guid("A04C572C-78BD-4B41-AE88-311D1DF8E743"));
                schemaBuilder.SetReadAccessLevel(AccessLevel.Public);//allow anyone to read the data
                schemaBuilder.SetWriteAccessLevel(AccessLevel.Public);//restrict writing to this vendor only
                schemaBuilder.SetVendorId("chliu"); //required because of write-access 
                schemaBuilder.SetSchemaName("OpeningRuleSet");
                //DataStorage createInfoStorage = DataStorage.Create(doc);
                //createInfoStorage.Name = "chLiuTest";

                //創建Field
                FieldBuilder fieldBuilder = schemaBuilder.AddSimpleField("fieldTest1", typeof(string));//create a filed to store an string
                //fieldBuilder.SetUnitType(UnitType.UT_Length);
                fieldBuilder.SetDocumentation("User data storage testing.");

                MessageBox.Show(dataToStore);
                Schema schema = schemaBuilder.Finish();//register the schema object
                Entity entity = new Entity(schema); //Create an entity (object) for this schema (class)
                Field fieldTest = schema.GetField("fieldTest1");
                entity.Set<string>(fieldTest, dataToStore);
                MessageBox.Show("資料儲存成功");

                ElementId id = new ElementId(1748581);
                Element elemPipe = doc.GetElement(id);
                elemPipe.SetEntity(entity); //store entity in the element
                //createInfoStorage.SetEntity(entity);

                //get data back
                Entity retrievedEntity = elemPipe.GetEntity(schema);
                string retrieveData = retrievedEntity.Get<string>(schema.GetField("fieldTest1"));
                MessageBox.Show(retrieveData);
                ;
                trans.Commit();
            }

            return Result.Succeeded;
        }
    }
}

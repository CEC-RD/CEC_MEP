using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace CEC_CADBlockTrans
{
    public class Availability : IExternalCommandAvailability
    {
        public bool IsCommandAvailable(UIApplication applicationData, CategorySet selectedCategories)
        {
            UIDocument uidoc = applicationData.ActiveUIDocument;
            if (uidoc.ActiveGraphicalView is ViewPlan)
            {
                return true;
            }
            return false;
        }
    }
}

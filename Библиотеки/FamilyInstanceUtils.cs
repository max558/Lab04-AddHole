using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitAPITreningLibrary
{
   public class FamilyInstanceUtils
    {
        public static FamilyInstance CreateFamilyInstance(
            ExternalCommandData commandData,
            FamilySymbol oFamSymb,
            XYZ insertionPoint,
            Level oLevel
            )
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            FamilyInstance familyInstance = null;

            //Вставка кземпляра семейства
            using (var ts=new Transaction(doc,"Create family instance"))
            {
                ts.Start();
                ActiveFamilySymbol(doc, oFamSymb);
                familyInstance = doc.Create.NewFamilyInstance(
                    insertionPoint,
                    oFamSymb,
                    oLevel,
                    StructuralType.NonStructural);

                ts.Commit();
            }

            return familyInstance;
        }

        /*
         *** === Активация семейства === ***
         * Входные данные:
         * doc - текущий досумент
         * familySymbol - семейство
         * Выходные данные:
         * Если семейство не является текущим, то становится таковым
         */
        public static void ActiveFamilySymbol(Document doc, FamilySymbol familySymbol)
        {
            using (Transaction ts = new Transaction(doc, "Активация семейства"))
            {
                if (!familySymbol.IsActive)
                {
                    familySymbol.Activate();
                    doc.Regenerate();
                }
            }
        }
    }
    
}

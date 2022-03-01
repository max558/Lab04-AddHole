using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using RevitAPITreningLibrary;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HolePlagin
{
    [Transaction(TransactionMode.Manual)]
    public class AddHole : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document arDoc = commandData.Application.ActiveUIDocument.Document;
            Document ovDoc = arDoc.Application.Documents.OfType<Document>().Where(x => x.Title.Contains("ОВ")).FirstOrDefault();

            if (ovDoc == null)
            {
                TaskDialog.Show("Ошибка", "Не найден связанный файл.");
                return Result.Cancelled;
            }

            FamilySymbol familySymbol = (SelectionUtils.SelectAllElementCategoryType<FamilySymbol>(arDoc, BuiltInCategory.OST_GenericModel))
                .Where(x => x.FamilyName.Equals("Прямоугольное отверстие")).FirstOrDefault();

            if (familySymbol == null)
            {
                TaskDialog.Show("Ошибка", "Семейство \"Прямоугольное отверстие\" не найдено.");
                return Result.Cancelled;
            }

            List<Duct> ducts = SelectionUtils.SelectAllElement<Duct>(ovDoc);
            List<Pipe> pipes = SelectionUtils.SelectAllElement<Pipe>(ovDoc);

            View3D view3D = SelectionUtils.SelectAllElement<View3D>(arDoc).Where(x => !x.IsTemplate).FirstOrDefault();

            if (view3D == null)
            {
                TaskDialog.Show("Ошибка", "Не найден 3D вмд.");
                return Result.Cancelled;
            }

            ReferenceIntersector referenceIntersector = new ReferenceIntersector(new ElementClassFilter(typeof(Wall)), FindReferenceTarget.Element, view3D);
            double widthEl = 4;
            double heightEl = 4;

            FamilyInstanceUtils.ActiveFamilySymbol(arDoc, familySymbol);

            //Отверстие по воздуховодам
            foreach (Duct duct in ducts)
            {
                //т.к. это общий метод для всех типов воздуховодов, то определение размеров выполняются в обработчике
                try
                {
                    widthEl = duct.Diameter;
                    heightEl = duct.Diameter;
                }
                catch (Autodesk.Revit.Exceptions.ApplicationException)
                {
                    try
                    {
                        widthEl = duct.Width;
                        heightEl = duct.Height;
                    }
                    catch (Autodesk.Revit.Exceptions.ApplicationException)
                    {
                    }
                }
                Line lineDuct = (duct.Location as LocationCurve).Curve as Line;
                AddHoleElement(arDoc, referenceIntersector, lineDuct, familySymbol, widthEl, heightEl);
            }

            widthEl = 4;
            heightEl = 4;
            // Отверстия по трубам
            foreach (Pipe pipe in pipes)
            {
                //т.к. это общий метод для всех типов воздуховодов и труб, то следующие команды выполняются в обработчике
                try
                {
                    widthEl = pipe.Diameter;
                    heightEl = pipe.Diameter;
                }
                catch (Autodesk.Revit.Exceptions.ApplicationException)
                {
                }
                Line linePipe = (pipe.Location as LocationCurve).Curve as Line;
                AddHoleElement(arDoc, referenceIntersector, linePipe, familySymbol, widthEl, heightEl);
            }


            return Result.Succeeded;
        }

        private void AddHoleElement(Document doc, ReferenceIntersector referenceIntersector, Line line, FamilySymbol familySymbol, double widthEl, double heightEl)
        {
            using (Transaction ts = new Transaction(doc, "Создание отверстий"))
            {
                ts.Start();

                XYZ point = line.GetEndPoint(0);
                XYZ direction = line.Direction;
                List<ReferenceWithContext> intersection = referenceIntersector.Find(point, direction)
                                                                              .Where(x => x.Proximity <= line.Length)
                                                                              .Distinct(new ReferenceWithContextElementEqualityComparer())
                                                                              .ToList();

                foreach (ReferenceWithContext refer in intersection)
                {
                    double proximity = refer.Proximity;
                    Reference reference = refer.GetReference();
                    Wall wall = doc.GetElement(reference.ElementId) as Wall;
                    Level level = doc.GetElement(wall.LevelId) as Level;
                    XYZ pointHole = point + (direction * proximity);

                    FamilyInstance hole = doc.Create.NewFamilyInstance(pointHole, familySymbol, wall, level, StructuralType.NonStructural);
                    //Настройка отверстий
                    Parameter width = hole.LookupParameter("Ширина");
                    Parameter height = hole.LookupParameter("Высота");
                    width.Set(widthEl);
                    height.Set(heightEl);
                }
                ts.Commit();
            }

        }
    }

    public class ReferenceWithContextElementEqualityComparer : IEqualityComparer<ReferenceWithContext>
    {
        public bool Equals(ReferenceWithContext x, ReferenceWithContext y)
        {
            if (ReferenceEquals(x, y)) return true;
            if (ReferenceEquals(null, x)) return false;
            if (ReferenceEquals(null, y)) return false;

            var xReference = x.GetReference();

            var yReference = y.GetReference();

            return xReference.LinkedElementId == yReference.LinkedElementId
                       && xReference.ElementId == yReference.ElementId;
        }

        public int GetHashCode(ReferenceWithContext obj)
        {
            var reference = obj.GetReference();

            unchecked
            {
                return (reference.LinkedElementId.GetHashCode() * 397) ^ reference.ElementId.GetHashCode();
            }
        }
    }
}

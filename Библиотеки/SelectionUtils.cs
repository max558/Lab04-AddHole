using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitAPITreningLibrary
{
    public class SelectionUtils
    {
        /*
         *** ==== Выбор всех типовых элементов модели по классу ==== ****
         * Входные параметры:
         * commandData - внешняя команда
         * Вывод: список элементов
         */
        public static List<T> SelectAllElement<T>(Document doc
            //ExternalCommandData commandData
            )
        {
            /*
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;
            */

            List<T> list = new FilteredElementCollector(doc)
                    .OfClass(typeof(T))
                    .OfType<T>()
                    //.Cast<T>()
                    .ToList();

            return list;
        }

        /*
         *** ==== Выбор всех типовых элементов модели по классу в активном представлении ==== ****
         * Входные параметры:
         * commandData - внешняя команда
         * viewId - ID представления. По умолчанию NULL - в активном документе
         * Вывод: список элементов
         */
        public static List<T> SelectAllElementActiveDoc<T>(
            Document doc,
            //ExternalCommandData commandData, 
            ElementId viewId = null)
        {
            /*
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;
            */

            List<T> list = new List<T>();
            if (viewId == null)
            {
                list = new FilteredElementCollector(doc, doc.ActiveView.Id)
                    .OfClass(typeof(T))
                    .OfType<T>()
                    //.Cast<T>()
                    .ToList();
            }
            else
            {
                list = new FilteredElementCollector(doc, viewId)
                .OfClass(typeof(T))
                .OfType<T>()
                //.Cast<T>()
                .ToList();
            }

            return list;
        }

        /*
         * Выбор всех экземпляров по категории
         */
        public static List<T> SelectAllElementCategory<T>(Document doc
            //ExternalCommandData commandData
            , BuiltInCategory builtInCategory)
        {
            /*
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;
            */

            List<T> list = new FilteredElementCollector(doc)
                .OfCategory(builtInCategory)
                .WhereElementIsNotElementType()
                .OfType<T>()
                //.Cast<T>()
                .ToList();

            return list;
        }

        /*
         * Выбор всех типов семейств по категории
         */
        public static List<T> SelectAllElementCategoryType<T>(Document doc
           // ExternalCommandData commandData
            , BuiltInCategory builtInCategory)
        {
            /*
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;
            */
           
            List<T> list = new FilteredElementCollector(doc)
                .OfCategory(builtInCategory)
                .WhereElementIsElementType()
                .OfType<T>()
                //.Cast<T>()
                .ToList();

            return list;
        }

        /*
         * Выбор элемента
         */
        public static Element PickObject(ExternalCommandData commandData, string message = "Выберите элемент:")
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            var selectedObject = uidoc.Selection.PickObject(ObjectType.Element, message);
            var oElement = doc.GetElement(selectedObject);
            return oElement;
        }

        /*
         * Множественный выбор элемента
         */
        public static List<Element> PickObjects(ExternalCommandData commandData, string message = "Выберите элементы:")
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            var selectedObjects = uidoc.Selection.PickObjects(ObjectType.Element, message);
            List<Element> elementList = new List<Element>();
            try
            {
                elementList = selectedObjects.Select(selectedObject => doc.GetElement(selectedObject)).ToList();
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            { }

            return elementList;
        }

        /*
         * Выбор объекта
         */
        public static T GetObject<T>(ExternalCommandData commandData, string message = "Выберите объект:")
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;
            
            Reference selectedObj = null;

            T elem;
            try
            {
                selectedObj = uidoc.Selection.PickObject(ObjectType.Element, message);
            }
            catch (Exception)
            {

                return default(T);
            }
            elem = (T)(object)doc.GetElement(selectedObj.ElementId);
            return elem;
        }

        /*
         * Выбор множества точек
         * message - Строковое сообщение
         * objectSnapTypes - выбор типа точек
         * counPoint - выбор количеста точек (0 - это выбор до отмены или ESC)
         */
        public static List<XYZ> GetPoints(ExternalCommandData commandData, string message, ObjectSnapTypes objectSnapTypes, int counPoint = 0)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;

            List<XYZ> points = new List<XYZ>();
            bool loopCondidtion = true;

            while (loopCondidtion)
            {
                XYZ pickedPoint = null;
                try
                {
                    pickedPoint = uidoc.Selection.PickPoint(objectSnapTypes, message);
                }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException ex)
                {
                    break;
                }
                points.Add(pickedPoint);
                //Проверка выбора количества точек
                if (counPoint > 0)
                {
                    loopCondidtion = !(points.Count == counPoint);
                }
            }

            return points;
        }
    }
}

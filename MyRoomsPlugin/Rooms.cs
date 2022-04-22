using Autodesk.Revit.Attributes;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyRoomsPlugin
{
    [TransactionAttribute(TransactionMode.Manual)]

    public class Rooms : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIApplication uiapp = commandData.Application;
                UIDocument uiDoc = commandData.Application.ActiveUIDocument;
                Document doc = uiDoc.Document;

                Level level = doc.ActiveView.GenLevel; //извлекаем уровень из текущего вида
                
                //далее размещаем помещения на активном виде
                Transaction transaction1 = new Transaction(doc);
                transaction1.Start("Размещение помещений");

                //чтобы вставить помещения на активный вид, нужно обратиться к документу, указать СВОЙСТВО "Create", дальше вызвать у него метод NewRooms2,
                //в качестве аргумента передаем уровень
                List <ElementId> roomListId = (List<ElementId>)doc.Create.NewRooms2(level);
               
                transaction1.Commit(); //метод commit подтверждает изменения

                //помещения размещены, теперь нужно расставить марки помещений
                //в семействе марки помещения есть параметр "Этаж", добавим в модели в категорию помещений параметр "Этаж" из ФОП
                //при помощи метода  CreateFloorParameter

                var categorySet = new CategorySet();
                categorySet.Insert(Category.GetCategory(doc, BuiltInCategory.OST_Rooms)); //находим категорию помещения

                Transaction transaction2 = new Transaction(doc);

                transaction2.Start("Добавление параметра");

                CreateFloorParameter(uiapp.Application, doc, "Этаж", categorySet, BuiltInParameterGroup.PG_DATA, true); //создаем параметр "Этаж"

                transaction2.Commit(); 

                //заполняем параметр "Этаж"
                
                Transaction transaction3 = new Transaction(doc);

                transaction3.Start("Запись параметра");
               
                List<Room> rooms = new FilteredElementCollector(doc)
                       .OfCategory(BuiltInCategory.OST_Rooms)
                       .OfType<Room>()
                       .ToList();
               
                foreach (var room in rooms)
                {
                    Parameter floorParam = room.LookupParameter("Этаж");
                    string floor = level.Name;
                    floorParam.Set(floor.Substring(floor.LastIndexOf("Level") + 6));
                }
               
                transaction3.Commit();

                Transaction transaction4 = new Transaction(doc);

                transaction4.Start("Добавление марок");

                FamilySymbol familySymbol = new FilteredElementCollector(doc) //находим семейство марки
                   .OfClass(typeof(FamilySymbol))
                   .OfCategory(BuiltInCategory.OST_RoomTags) //марки помещений
                   .OfType<FamilySymbol>()
                   .Where(x => x.FamilyName.Equals("Марка помещения Этаж_Номер"))
                   .FirstOrDefault();

                if (familySymbol == null) //проверяем, что семейство отверстия загружено в модель
                {
                    TaskDialog.Show("Ошибка", "Не найдено семейство \"Марка помещения\"");
                    return Result.Cancelled;
                }

                foreach (var room in rooms)
                {
                    // находим центр комнаты
                    Element element = room;
                    XYZ roomCenter = GetElementCenter(room);
                    IndependentTag.Create(doc, familySymbol.Id, doc.ActiveView.Id, new Reference(room), false,
                   TagOrientation.Horizontal, roomCenter);
                }

                transaction4.Commit();

            }
            catch (Exception ex)
            {
                message = ex.Message; 
                return Result.Failed;
            }

            return Result.Succeeded;
        }

        private XYZ GetElementCenter(Element element) //метод для нахождения центра элемента с использованием метода get_BoundingBox
        {
            BoundingBoxXYZ bounding = element.get_BoundingBox(null);
            return (bounding.Max + bounding.Min) / 2;
        }

        //метод создания параметра "Этаж"
        private void CreateFloorParameter(Application application, Document doc, string parameterName,
                                         CategorySet categorySet, BuiltInParameterGroup builtInParameterGroup, bool isInstance)
        {
            DefinitionFile definitionFile = application.OpenSharedParameterFile(); //открываем файл общих параметров
            if (definitionFile == null)
            {
                TaskDialog.Show("Ошибка", "Не найден файл общих параметров");
                return;
            }

            Definition definition = definitionFile.Groups //открываем файл общих параметров
                      .SelectMany(group => group.Definitions) //из всех групп выбираем все определения параметров
                      .FirstOrDefault(def => def.Name.Equals(parameterName));
            
            //при помощи метода FirstOrDefault выбрали только один параметр с заранее заданным именем

            if (definition == null)
            {
                TaskDialog.Show("Ошибка", "Не найден указанный параметр");
                return;
            }

            Binding binding = application.Create.NewTypeBinding(categorySet); 
            if (isInstance) 
                binding = application.Create.NewInstanceBinding(categorySet); //создаем параметр экземпляра

            BindingMap map = doc.ParameterBindings; //создаем переменную для вставки параметра в проект
            map.Insert(definition, binding, builtInParameterGroup); //добавляем параметр, перечисляя аргументы (параметр, тип или экземпляр, группа)
        }
    }
}



using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RoomNumbe
{
    [Transaction(TransactionMode.Manual)]
    public class NewRoom : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;  //создаем документ

            List<Level> levels = new FilteredElementCollector(doc)
                .OfClass(typeof(Level))  //фильтр по этажам
                .OfType<Level>()
                .ToList();

            FamilySymbol familySymbol = GetFamilySymbol(doc, "Марка помещения");

            if (familySymbol == null)  //проверка
            {
                TaskDialog.Show("Ошибка", "Не найдено семейство \"Марка помещения\"");
                return Result.Cancelled;
            }

            View view = doc.ActiveView;

            List<Room> rooms = new FilteredElementCollector(doc)  //фильтр помещений
                .OfCategory(BuiltInCategory.OST_Rooms)
                .OfType<Room>()
                .ToList();

            if (rooms.Count != 0)  //проверка
            {
                TaskDialog td = new TaskDialog("Внимание");
                td.MainContent = "В модели обнаружены помещения. Удалить и продолжить?";
                td.CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No;
                TaskDialogResult tdr = td.Show();
                if (tdr == TaskDialogResult.Yes)
                {
                    Transaction tr1 = new Transaction(doc, "Delete Existing Rooms");
                    tr1.Start();
                    foreach (Room room in rooms)
                    {
                        doc.Delete(room.Id);
                    }
                    tr1.Commit();
                }
                else if (tdr == TaskDialogResult.No)
                {
                    return Result.Cancelled;
                }
                else
                {
                    return Result.Failed;
                }
            }

            Transaction transaction = new Transaction(doc, "Create Rooms and Room Tags");  //транзакция
            transaction.Start();
            int lev = 1;

            foreach (Level level in levels)
            {
                PlanTopology planTopology = null;
                planTopology = doc.get_PlanTopology(level);
                int num = 1;

                foreach (PlanCircuit circuit in planTopology.Circuits)
                {
                    Room room = null;
                    room = doc.Create.NewRoom(null, circuit);
                    room.Name = "Помещение " + num;
                    room.Number = lev + "_" + num;
                    LinkElementId link = new LinkElementId(room.Id);
                    XYZ insertPoint = GetElementCenter(room);
                    UV point = new UV(insertPoint.X, insertPoint.Y);
                    doc.Create.NewRoomTag(link, point, view.Id);
                    num++;
                }
                lev++;
            }

            transaction.Commit();

            return Result.Succeeded;
        }

        private FamilySymbol GetFamilySymbol(Document doc, string str)
        {
            FamilySymbol familySymbol = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_RoomTags)
                .OfType<FamilySymbol>()
                .Where(x => x.FamilyName.Equals(str))
                .FirstOrDefault();
            return familySymbol;
        }

        private XYZ GetElementCenter(Element element)
        {
            BoundingBoxXYZ bounding = element.get_BoundingBox(null);
            XYZ elementCenter = (bounding.Max + bounding.Min) / 2;
            return elementCenter;
        }
    }
}

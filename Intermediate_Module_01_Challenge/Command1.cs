#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;

#endregion

namespace Intermediate_Module_01_Challenge
{
    [Transaction(TransactionMode.Manual)]
    public class Command1 : IExternalCommand
    {
        public object levelName { get; private set; }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;

            // Your code goes here
            using (Transaction t = new Transaction(doc))
            {
                t.Start("Create Schedule");
                ElementId catId = new ElementId(BuiltInCategory.OST_Rooms);
                ViewSchedule newSchedule = ViewSchedule.CreateSchedule(doc, catId);
                newSchedule.Name = "DEPT - BUILDING SERVICES";

                FilteredElementCollector roomCollector = new FilteredElementCollector(doc);
                roomCollector.OfCategory(BuiltInCategory.OST_Rooms);
                roomCollector.WhereElementIsNotElementType();

                Element roomInst = roomCollector.FirstElement();
                Parameter roomNumParam = roomInst.LookupParameter("Number");
                Parameter roomNameParam = roomInst.LookupParameter("Name");
                Parameter roomDepartParam = roomInst.LookupParameter("Department");
                Parameter roomCommentsParam = roomInst.LookupParameter("Comments");
                Parameter roomAreaParam = roomInst.LookupParameter("Area");
                Parameter roomLevelParam = roomInst.LookupParameter("Level");

                ScheduleField roomNumField =
                    newSchedule.Definition.AddField(ScheduleFieldType.Instance, roomNumParam.Id);
                ScheduleField roomNameField =
                    newSchedule.Definition.AddField(ScheduleFieldType.Instance, roomNameParam.Id);
                ScheduleField roomDeptField =
                    newSchedule.Definition.AddField(ScheduleFieldType.Instance, roomDepartParam.Id);
                ScheduleField roomCommentsField =
                    newSchedule.Definition.AddField(ScheduleFieldType.Instance, roomCommentsParam.Id);
                ScheduleField roomAreaField =
                    newSchedule.Definition.AddField(ScheduleFieldType.ViewBased, roomAreaParam.Id);
                ScheduleField roomLevelField =
                    newSchedule.Definition.AddField(ScheduleFieldType.Instance, roomLevelParam.Id);

                Level filterLevel = GetLevelByName(doc, levelName);
                ScheduleFilter levelFilter = new ScheduleFilter(roomLevelField.FieldId, ScheduleFilterType.Equal, filterLevel.Id);
                newSchedule.Definition.AddFilter(levelFilter);

                ScheduleSortGroupField levelSort = new ScheduleSortGroupField(roomLevelField.FieldId);
                newSchedule.Definition.AddSortGroupField(levelSort);
                

                ScheduleSortGroupField roomNameSort = new ScheduleSortGroupField( roomNameField.FieldId);
                newSchedule.Definition.AddSortGroupField(roomNameSort);

                newSchedule.Definition.IsItemized = true;
                

                t.Commit();
            }

            return Result.Succeeded;
        }

        private Level GetLevelByName(Document doc, object levelName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfCategory(BuiltInCategory.OST_Levels);
            collector.WhereElementIsNotElementType();

            foreach (Level curLevel in collector)
            {
                if (curLevel.Name == levelName)
                    return curLevel;
            }
            return null;
        }

        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "btnCommand1";
            string buttonTitle = "Button 1";

            ButtonDataClass myButtonData1 = new ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Blue_32,
                Properties.Resources.Blue_16,
                "This is a tooltip for Button 1");

            return myButtonData1.Data;
        }
    }
}

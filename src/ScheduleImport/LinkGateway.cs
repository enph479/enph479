using System;
using Autodesk.Revit.DB.ExtensibleStorage;
using Autodesk.Revit.DB.Electrical;

namespace ElectricalToolSuite.ScheduleImport
{
    static class LinkGateway
    {
        private const string SchemaName = "ExcelScheduleLinkSchema";
        private const string WorkbookPathKey = "WorkbookPath";
        private const string WorksheetNameKey = "WorksheetName";
        private const string ScheduleTypeKey = "ScheduleType";

        private static readonly Guid SchemaGuid = new Guid("79d73fc5-87e9-4681-9508-6d9bc27b0ec8");

        private static Schema GetSchema()
        {
            var schema = Schema.Lookup(SchemaGuid);
            if (schema != null)
                return schema;

            var schemaBuilder = new SchemaBuilder(SchemaGuid);

            schemaBuilder.SetReadAccessLevel(AccessLevel.Public);
            schemaBuilder.SetWriteAccessLevel(AccessLevel.Public);
            schemaBuilder.SetSchemaName(SchemaName);

            var fieldBuilder = schemaBuilder.AddSimpleField(WorkbookPathKey, typeof(string));
            fieldBuilder.SetDocumentation("The path to the Excel workbook from which this schedule was created.");

            fieldBuilder = schemaBuilder.AddSimpleField(WorksheetNameKey, typeof(string));
            fieldBuilder.SetDocumentation("The name of the Excel worksheet within the workbook from which this schedule was created.");

            fieldBuilder = schemaBuilder.AddSimpleField(ScheduleTypeKey, typeof(string));
            fieldBuilder.SetDocumentation("The type of this schedule.");

            return schemaBuilder.Finish();
        }

        public static bool IsLinked(PanelScheduleView schedule)
        {
            return schedule.GetEntity(GetSchema()).IsValid();
        }

        public static string GetWorkbookPath(PanelScheduleView schedule)
        {
            return schedule.GetEntity(GetSchema()).Get<string>(WorkbookPathKey);
        }

        public static string GetWorksheetName(PanelScheduleView schedule)
        {
            return schedule.GetEntity(GetSchema()).Get<string>(WorksheetNameKey);
        }

        public static string GetScheduleType(PanelScheduleView schedule)
        {
            return schedule.GetEntity(GetSchema()).Get<string>(ScheduleTypeKey);
        }

        public static void SetWorkbookPath(PanelScheduleView schedule, string workbookPath)
        {
            var ent = schedule.GetEntity(GetSchema());
            ent.Set(WorkbookPathKey, workbookPath);
            schedule.SetEntity(ent);
        }

        public static void SetWorksheetName(PanelScheduleView schedule, string worksheetName)
        {
            var ent = schedule.GetEntity(GetSchema());
            ent.Set(WorksheetNameKey, worksheetName);
            schedule.SetEntity(ent);
        }

        public static void SetScheduleType(PanelScheduleView schedule, string scheduleType)
        {
            var ent = schedule.GetEntity(GetSchema());
            ent.Set(ScheduleTypeKey, scheduleType);
            schedule.SetEntity(ent);
        }

        public static void CreateLink(PanelScheduleView schedule, string workbookPath, string worksheetName,
            string scheduleType)
        {
            if (IsLinked(schedule))
                throw new ArgumentException("Cannot create link because schedule is already linked", "schedule");

            var schema = GetSchema();

            var entity = new Entity(schema);

            var pathField = schema.GetField(WorkbookPathKey);
            entity.Set(pathField, workbookPath);

            var sheetField = schema.GetField(WorksheetNameKey);
            entity.Set(sheetField, worksheetName);

            var typeField = schema.GetField(ScheduleTypeKey);
            entity.Set(typeField, scheduleType);

            schedule.SetEntity(entity);
        }

        public static void DeleteLink(PanelScheduleView schedule)
        {
            if (!IsLinked(schedule))
                throw new ArgumentException("Cannot delete link because schedule is not linked", "schedule");

            if (!schedule.DeleteEntity(GetSchema()))
                throw new ApplicationException("Failed to delete linked schedule schema");
        }
    }
}

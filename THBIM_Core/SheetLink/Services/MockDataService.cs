using System.Collections.Generic;
using THBIM.Models;

namespace THBIM.Services
{
    public class MockDataService
    {
        public List<CategoryItem> GetModelCategories() => new()
        {
            new("Analytical Members",   "Structure"),
            new("Ceilings",             "Architecture"),
            new("Columns",              "Architecture"),
            new("Curtain Panels",       "Architecture"),
            new("Curtain Systems",      "Architecture"),
            new("Doors",                "Architecture"),
            new("Floors",               "Architecture"),
            new("Furniture",            "Architecture"),
            new("HVAC Zones",           "MEP"),
            new("Levels",               "Architecture"),
            new("Pipe Accessories",     "MEP"),
            new("Pipes",                "MEP"),
            new("Piping Systems",       "MEP"),
            new("Roofs",                "Architecture"),
            new("Rooms",                "Architecture"),
            new("Stairs",               "Architecture"),
            new("Structural Columns",   "Structure"),
            new("Structural Framing",   "Structure"),
            new("Walls",                "Architecture"),
            new("Windows",              "Architecture"),
        };

        public List<CategoryItem> GetAnnotationCategories() => new()
        {
            new("Callout Heads",           "Architecture"),
            new("Dimensions",              "Architecture"),
            new("Elevation Marks",         "Architecture"),
            new("Generic Annotations",     "Architecture"),
            new("Grid Heads",              "Architecture"),
            new("Grids",                   "Architecture"),
            new("Level Heads",             "Architecture"),
            new("Reference Planes",        "Architecture"),
            new("Room Tags",               "Architecture"),
            new("Section Marks",           "Architecture"),
            new("Text Notes",              "Architecture"),
            new("Title Blocks",            "Architecture"),
        };

        public List<ParameterItem> GetParameters(string categoryName) => new()
        {
            new("Area",                 "Double",    ParamKind.ReadOnly),
            new("Assembly Code",        "String",    ParamKind.Type),
            new("Base Constraint",      "ElementId", ParamKind.Instance),
            new("Base Finish",          "String",    ParamKind.Instance),
            new("Base Offset",          "Double",    ParamKind.Instance),
            new("Category",             "String",    ParamKind.ReadOnly),
            new("Ceiling Finish",       "String",    ParamKind.Instance),
            new("Comments",             "String",    ParamKind.Instance),
            new("Department",           "String",    ParamKind.Instance),
            new("Description",          "String",    ParamKind.Type),
            new("Family",               "String",    ParamKind.ReadOnly),
            new("Family and Type",      "String",    ParamKind.ReadOnly),
            new("Fire Rating",          "String",    ParamKind.Type),
            new("Floor Finish",         "String",    ParamKind.Instance),
            new("Height",               "Double",    ParamKind.Instance),
            new("Level",                "ElementId", ParamKind.Instance),
            new("Mark",                 "String",    ParamKind.Instance),
            new("Occupancy",            "String",    ParamKind.Instance),
            new("Perimeter",            "Double",    ParamKind.ReadOnly),
            new("Phase Created",        "ElementId", ParamKind.Instance),
            new("Type Comments",        "String",    ParamKind.Type),
            new("Type Mark",            "String",    ParamKind.Type),
            new("Volume",               "Double",    ParamKind.ReadOnly),
            new("Wall Finish",          "String",    ParamKind.Instance),
            new("Width",                "Double",    ParamKind.Type),
        };

        public List<CategoryItem> GetElements(string categoryName) => new()
        {
            new($"{categoryName} - Type 01", categoryName),
            new($"{categoryName} - Type 02", categoryName),
            new($"{categoryName} - Type 03", categoryName),
        };

        public List<ScheduleItem> GetSchedules() => new()
        {
            new("Door Schedule",                       ScheduleKind.Regular),
            new("Electrical Analytical Bus Schedule",  ScheduleKind.Regular),
            new("Electrical Analytical Load Schedule", ScheduleKind.Regular),
            new("Level Schedule",                      ScheduleKind.Regular),
            new("Room Finish Schedule",                ScheduleKind.Regular),
            new("Sheet Index",                         ScheduleKind.SheetList),
            new("Space Outdoor Air Schedule",          ScheduleKind.Regular),
            new("Structural Column Schedule",          ScheduleKind.Regular),
            new("Structural Foundation Schedule",      ScheduleKind.Regular),
            new("Window Schedule",                     ScheduleKind.Regular),
            new("Working_Sheet List",                  ScheduleKind.SheetList),
            new("Working_View List",                   ScheduleKind.ViewList),
        };

        public List<SpatialItem> GetRooms() => new()
        {
            new(101001, "101", "Office A",        "New Construction"),
            new(101002, "102", "Office B",        "New Construction"),
            new(101003, "103", "Conference Room", "New Construction"),
            new(101004, "201", "Open Space",      "New Construction"),
            new(101005, "202", "Break Room",      "New Construction"),
            new(101006, "203", "Storage",         "New Construction"),
            new(101007, "301", "Lobby",           "New Construction"),
            new(101008, "302", "Corridor 01",     "New Construction"),
            new(101009, "303", "WC Male",         "New Construction"),
            new(101010, "304", "WC Female",       "New Construction"),
        };

        public List<SpatialItem> GetSpaces() => new()
        {
            new(102001, "S101", "Office A (Space)",    "New Construction", false),
            new(102002, "S102", "Office B (Space)",    "New Construction", false),
            new(102003, "S103", "Corridor (Space)",    "New Construction", false),
            new(102004, "S201", "Open Space (Space)",  "New Construction", false),
        };

        public List<ParameterItem> GetSpatialParameters() => new()
        {
            new("Area",           "Double",    ParamKind.ReadOnly),
            new("Base Finish",    "String",    ParamKind.Instance),
            new("Ceiling Finish", "String",    ParamKind.Instance),
            new("Comments",       "String",    ParamKind.Instance),
            new("Department",     "String",    ParamKind.Instance),
            new("Floor Finish",   "String",    ParamKind.Instance),
            new("Name",           "String",    ParamKind.Instance),
            new("Number",         "String",    ParamKind.Instance),
            new("Occupancy",      "String",    ParamKind.Instance),
            new("Perimeter",      "Double",    ParamKind.ReadOnly),
            new("Volume",         "Double",    ParamKind.ReadOnly),
            new("Wall Finish",    "String",    ParamKind.Instance),
        };
    }
}

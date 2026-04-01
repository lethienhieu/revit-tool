using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Mechanical;
using System;
using System.Collections.Generic;
using System.Linq;

namespace THBIM.Helpers
{
    public static class QTOHelper
    {
        public static List<string> GetUsedSystemTypeNames(Document doc)
        {
            return new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_PipeCurves)
                .WhereElementIsNotElementType()
                .Cast<Pipe>()
                .Select(p => p.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM)?.AsValueString())
                .Where(name => !string.IsNullOrEmpty(name))
                .Select(name => name!) // ✅ Fix CS8619: Khẳng định giá trị không null
                .Distinct()
                .ToList();
        }

        public static List<string> GetUsedDuctSystemTypeNames(Document doc)
        {
            return new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_DuctCurves)
                .WhereElementIsNotElementType()
                .Cast<Duct>()
                .Select(d => d.get_Parameter(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM)?.AsValueString())
                .Where(name => !string.IsNullOrEmpty(name))
                .Select(name => name!) // ✅ Fix CS8619
                .Distinct()
                .ToList();
        }

        public static List<QTO_Pipe> GetPipeInfo(Document doc, List<string> systems, LengthUnit unit)
        {
            // ✅ Fix CS8604: Dùng (?? string.Empty) để tránh truyền null vào Contains
            var pipes = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_PipeCurves)
                .WhereElementIsNotElementType()
                .Cast<Pipe>()
                .Where(p => systems.Contains(p.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM)?.AsValueString() ?? string.Empty))
                .Cast<Element>();

            var flexPipes = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_FlexPipeCurves)
                .WhereElementIsNotElementType()
                .Cast<FlexPipe>()
                .Where(fp => systems.Contains(fp.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM)?.AsValueString() ?? string.Empty))
                .Cast<Element>();

            var allPipes = pipes.Concat(flexPipes);

            return allPipes
                .GroupBy(p => new
                {
                    System = p.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM)?.AsValueString(),
                    Type = p.Name,
                    Size = p.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)?.AsValueString()
                })
                .Select(g => new QTO_Pipe
                {
                    SystemName = g.Key.System,
                    TypeName = g.Key.Type,
                    Diameter = g.Key.Size,
                    Length = FormatLength(g.Sum(p => p.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble()), unit),
                    ElementId = g.First().Id,
                    Reference = new Reference(g.First())
                })
                .ToList();
        }

        public static List<QTO_Pipe> GetDuctInfo(Document doc, List<string> systems, LengthUnit unit)
        {
            var ducts = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_DuctCurves)
                .WhereElementIsNotElementType()
                .Cast<Duct>()
                .Where(d => systems.Contains(d.get_Parameter(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM)?.AsValueString() ?? string.Empty))
                .Cast<Element>();

            var flexDucts = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_FlexDuctCurves)
                .WhereElementIsNotElementType()
                .Cast<FlexDuct>()
                .Where(fd => systems.Contains(fd.get_Parameter(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM)?.AsValueString() ?? string.Empty))
                .Cast<Element>();

            var allDucts = ducts.Concat(flexDucts);

            return allDucts
                .GroupBy(d => new
                {
                    System = d.get_Parameter(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM)?.AsValueString(),
                    Type = d.Name,
                    Size = d is Duct duct
                        ? GetDuctSizeString(duct)
                        : d is FlexDuct flex
                            ? flex.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)?.AsValueString()
                            : "N/A"
                })
                .Select(g => new QTO_Pipe
                {
                    SystemName = g.Key.System,
                    TypeName = g.Key.Type,
                    Diameter = g.Key.Size,
                    Length = FormatLength(g.Sum(d => d.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble()), unit),
                    ElementId = g.First().Id,
                    Reference = new Reference(g.First())
                })
                .ToList();
        }

        public static List<QTO_Fitting> GetFittingInfo(Document doc, List<string> systems)
        {
            return new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_PipeFitting)
                .WhereElementIsNotElementType()
                .Cast<FamilyInstance>()
                .Where(f => systems.Contains(f.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM)?.AsValueString() ?? string.Empty))
                .GroupBy(f => new
                {
                    Type = f.Name,
                    Size = f.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE)?.AsValueString(),
                    CategoryId = f.Category?.Id.GetValue(),
                    Family = f.Symbol?.Family?.Name
                })
                .Select(g =>
                {
                    string category = (g.Key.CategoryId == (long)BuiltInCategory.OST_PipeAccessory) ? "Accessory" : "Fitting";

                    return new QTO_Fitting
                    {
                        FamilyName = g.Key.Family,
                        TypeName = g.Key.Type,
                        Size = g.Key.Size,
                        Count = g.Count(),
                        Category = category,
                        ElementId = g.First().Id,
                        Reference = new Reference(g.First())
                    };
                })
                .ToList();
        }

        [Obsolete]
        public static List<QTO_Fitting> GetDuctFittingInfo(Document doc, List<string> systems)
        {
            return new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .Where(e =>
                    e.Category != null &&
                    (e.Category.Id.GetValue() == (long)BuiltInCategory.OST_DuctFitting ||
                     e.Category.Id.GetValue() == (long)BuiltInCategory.OST_DuctAccessory))
                .Cast<FamilyInstance>()
                .Where(f => systems.Contains(f.get_Parameter(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM)?.AsValueString() ?? string.Empty))
                .GroupBy(f => new
                {
                    Type = f.Name,
                    Size = f.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE)?.AsValueString(),
                    CategoryId = f.Category?.Id.GetValue(),
                    Family = f.Symbol?.Family?.Name
                })
                .Select(g =>
                {
                    string category = (g.Key.CategoryId == (long)BuiltInCategory.OST_DuctAccessory) ? "Accessory" : "Fitting";

                    return new QTO_Fitting
                    {
                        FamilyName = g.Key.Family,
                        TypeName = g.Key.Type,
                        Size = g.Key.Size,
                        Count = g.Count(),
                        Category = category,
                        ElementId = g.First().Id,
                        Reference = new Reference(g.First())
                    };
                })
                .ToList();
        }

        public static List<QTO_Insulation> GetInsulationInfo(Document doc, List<string> systems, LengthUnit unit)
        {
            return new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_PipeInsulations)
                .WhereElementIsNotElementType()
                .Cast<InsulationLiningBase>()
                .Where(i =>
                {
                    var host = doc.GetElement(i.HostElementId);
                    // ✅ Fix CS8604
                    return host is Pipe && systems.Contains(
                        host.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM)?.AsValueString() ?? string.Empty);
                })
                .Select(i =>
                {
                    var host = doc.GetElement(i.HostElementId) as Pipe;
                    var system = host?.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM)?.AsValueString();
                    var size = host?.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER)?.AsValueString();
                    var thickness = i.LookupParameter("Insulation Thickness")?.AsValueString();
                    var length = i.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble();

                    return new QTO_Insulation
                    {
                        SystemName = system,
                        Diameter = size,
                        InsulationThickness = thickness,
                        LengthRaw = length
                    };
                })
                .GroupBy(i => new { i.SystemName, i.Diameter, i.InsulationThickness })
                .OrderBy(g => g.Key.Diameter)
                .ThenBy(g => g.Key.InsulationThickness)
                .Select(g => new QTO_Insulation
                {
                    SystemName = g.Key.SystemName,
                    Diameter = g.Key.Diameter,
                    InsulationThickness = g.Key.InsulationThickness,
                    Length = FormatLength(g.Sum(x => x.LengthRaw), unit),
                })
                .ToList();
        }

        public static List<QTO_Insulation> GetDuctInsulationInfo(Document doc, List<string> systems, LengthUnit unit)
        {
            var ductIns = new FilteredElementCollector(doc)
                .OfClass(typeof(DuctInsulation))
                .Cast<InsulationLiningBase>()
                .Where(i =>
                {
                    var host = doc.GetElement(i.HostElementId);
                    bool isValidHost = host is Duct || host is FlexDuct;
                    // ✅ Fix CS8604
                    return isValidHost &&
                           systems.Contains(host.get_Parameter(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM)?.AsValueString() ?? string.Empty);
                })
                .Select(i =>
                {
                    var host = doc.GetElement(i.HostElementId) as Duct;
                    var system = host?.get_Parameter(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM)?.AsValueString();
                    var size = host is Duct duct ? GetDuctSizeString(duct) : "N/A";
                    var thickness = i.LookupParameter("Insulation Thickness")?.AsValueString();
                    var length = i.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble();

                    return new QTO_Insulation
                    {
                        SystemName = system,
                        Diameter = size,
                        InsulationThickness = thickness,
                        LengthRaw = length,
                    };
                });

            return ductIns
                .GroupBy(i => new { i.SystemName, i.Diameter, i.InsulationThickness })
                .OrderBy(g => g.Key.Diameter)
                .ThenBy(g => g.Key.InsulationThickness)
                .Select(g => new QTO_Insulation
                {
                    SystemName = g.Key.SystemName,
                    Diameter = g.Key.Diameter,
                    InsulationThickness = g.Key.InsulationThickness,
                    Length = FormatLength(g.Sum(x => x.LengthRaw), unit)
                })
                .ToList();
        }

        public static List<QTO_Pipe> GetPipeInfoFromElements(Document doc, List<Element> elements, LengthUnit unit)
        {
            var pipes = elements
                .Where(e => e is Pipe || e is FlexPipe)
                .Cast<Element>();

            return pipes
                .GroupBy(p => new
                {
                    System = p.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM)?.AsValueString(),
                    Type = p.Name,
                    Size = p.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)?.AsValueString()
                })
                .Select(g => new QTO_Pipe
                {
                    SystemName = g.Key.System,
                    TypeName = g.Key.Type,
                    Diameter = g.Key.Size,
                    Length = FormatLength(g.Sum(p => p.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble()), unit)
                })
                .ToList();
        }

        public static List<QTO_Pipe> GetDuctInfoFromElements(Document doc, List<Element> elements, LengthUnit unit)
        {
            var ducts = elements
                .Where(e => e is Duct || e is FlexDuct)
                .Cast<Element>();

            return ducts
                .GroupBy(d => new
                {
                    System = d.get_Parameter(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM)?.AsValueString(),
                    Type = d.Name,
                    Size = d is Duct duct
                    ? GetDuctSizeString(duct)
                    : d is FlexDuct flex
                    ? flex.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)?.AsValueString()
                    : "N/A"
                })
                .Select(g => new QTO_Pipe
                {
                    SystemName = g.Key.System,
                    TypeName = g.Key.Type,
                    Diameter = g.Key.Size,
                    Length = FormatLength(g.Sum(d => d.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble()), unit)
                })
                .ToList();
        }

        public static List<QTO_Fitting> GetFittingInfoFromElements(Document doc, List<Element> elements)
        {
            return elements
                .Where(e => e.Category != null &&
                            (e.Category.Id.GetValue() == (long)BuiltInCategory.OST_PipeFitting ||
                             e.Category.Id.GetValue() == (long)BuiltInCategory.OST_PipeAccessory ||
                             e.Category.Id.GetValue() == (long)BuiltInCategory.OST_DuctFitting ||
                             e.Category.Id.GetValue() == (long)BuiltInCategory.OST_DuctAccessory))
                .OfType<FamilyInstance>()
                .GroupBy(f => new
                {
                    Type = f.Name,
                    Size = f.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE)?.AsValueString(),
                    CategoryId = f.Category?.Id.GetValue(),
                    Family = f.Symbol?.Family?.Name
                })
                .Select(g =>
                {
                    string category = (g.Key.CategoryId == (long)BuiltInCategory.OST_PipeAccessory) ? "Accessory" : "Fitting";

                    return new QTO_Fitting
                    {
                        FamilyName = g.Key.Family,
                        TypeName = g.Key.Type,
                        Size = g.Key.Size,
                        Count = g.Count(),
                        Category = category
                    };
                })
                .ToList();
        }

        public static List<QTO_Insulation> GetInsulationInfoFromElements(Document doc, List<Element> elements, LengthUnit unit)
        {
            return elements
                .OfType<InsulationLiningBase>()
                .Select(i =>
                {
                    var host = doc.GetElement(i.HostElementId);
                    if (!(host is Pipe || host is Duct || host is FlexDuct))
                        return null;

                    string? systemName = host.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM)?.AsValueString()
                                     ?? host.get_Parameter(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM)?.AsValueString();

                    string? size = "N/A";
                    if (host is Pipe pipe)
                        size = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER)?.AsValueString();
                    else if (host is Duct duct)
                        size = GetDuctSizeString(duct);
                    else if (host is FlexDuct flex)
                        size = flex.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)?.AsValueString();

                    string? thickness = i.LookupParameter("Insulation Thickness")?.AsValueString();
                    double length = i.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble();

                    return new QTO_Insulation
                    {
                        SystemName = systemName,
                        Diameter = size,
                        InsulationThickness = thickness,
                        LengthRaw = length
                    };
                })
                .Where(i => i != null)
                .Select(i => i!) // Fix CS8619 cho list
                .GroupBy(i => new { i.SystemName, i.Diameter, i.InsulationThickness })
                .OrderBy(g => g.Key.Diameter)
                .ThenBy(g => g.Key.InsulationThickness)
                .Select(g => new QTO_Insulation
                {
                    SystemName = g.Key.SystemName,
                    Diameter = g.Key.Diameter,
                    InsulationThickness = g.Key.InsulationThickness,
                    Length = FormatLength(g.Sum(x => x.LengthRaw), unit)
                })
                .ToList();
        }

        private static string FormatLength(double lengthInFeet, LengthUnit unit)
        {
            double value = unit switch
            {
                LengthUnit.Meter => UnitUtils.ConvertFromInternalUnits(lengthInFeet, UnitTypeId.Meters),
                LengthUnit.Inch => UnitUtils.ConvertFromInternalUnits(lengthInFeet, UnitTypeId.Inches),
                _ => UnitUtils.ConvertFromInternalUnits(lengthInFeet, UnitTypeId.Millimeters),
            };
            return Math.Round(value, 2).ToString();
        }

        private static string GetDuctSizeString(Duct duct)
        {
            double width = duct.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM)?.AsDouble() ?? 0;
            double height = duct.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)?.AsDouble() ?? 0;

            if (width > 0 && height > 0)
            {
                width = UnitUtils.ConvertFromInternalUnits(width, UnitTypeId.Millimeters);
                height = UnitUtils.ConvertFromInternalUnits(height, UnitTypeId.Millimeters);
                return $"{Math.Round(width)} x {Math.Round(height)}";
            }
            var calcSize = duct.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE);
            return calcSize?.AsValueString() ?? "N/A";
        }
    }

    // ✅ Fix CS8618: Thêm dấu ? để cho phép thuộc tính có thể null
    public class QTO_Pipe
    {
        public string? SystemName { get; set; }
        public string? TypeName { get; set; }
        public string? Diameter { get; set; }
        public string? Length { get; set; }
        public string? Size => Diameter;
        public double LengthRaw { get; set; }
        public ElementId? ElementId { get; set; }
        public Reference? Reference { get; set; }
    }

    public class QTO_Fitting
    {
        public string? FamilyName { get; set; }
        public string? TypeName { get; set; }
        public string? Size { get; set; }
        public int Count { get; set; }
        public string? Category { get; set; }
        public ElementId? ElementId { get; set; }
        public Reference? Reference { get; set; }
    }

    public class QTO_Insulation
    {
        public string? SystemName { get; set; }
        public string? Diameter { get; set; }
        public string? Length { get; set; }
        public string? InsulationThickness { get; set; }
        public double LengthRaw { get; set; }
        public ElementId? ElementId { get; set; }
        public Reference? Reference { get; set; }
    }

    public enum LengthUnit
    {
        Millimeter,
        Meter,
        Inch
    }
}
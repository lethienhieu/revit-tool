using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;

namespace THBIM.MEP.Core;

internal static class ConnectorUtils
{
    public static bool IsUsable(Connector connector)
        => connector.Domain != Domain.DomainUndefined
           && connector.ConnectorType != ConnectorType.Logical;

    public static bool IsOpen(Connector connector)
        => IsUsable(connector) && !connector.IsConnected;

    public static ConnectorSet? GetConnectors(Element element)
        => element switch
        {
            MEPCurve mepCurve => mepCurve.ConnectorManager?.Connectors,
            FamilyInstance familyInstance => familyInstance.MEPModel?.ConnectorManager?.Connectors,
            FabricationPart fabricationPart => fabricationPart.ConnectorManager?.Connectors,
            _ => null
        };

    public static IList<Connector> ToList(Element element)
        => GetConnectors(element)?.Cast<Connector>().ToList() ?? [];

    public static IList<Connector> GetOpenConnectors(Element element)
        => ToList(element).Where(IsOpen).ToList();

    public static Connector? GetClosestConnector(Element element, XYZ point)
        => ToList(element)
            .Where(IsUsable)
            .OrderBy(c => c.Origin.DistanceTo(point))
            .FirstOrDefault();

    public static Connector? GetClosestOpenConnector(Element element, XYZ point)
        => GetOpenConnectors(element)
            .OrderBy(c => c.Origin.DistanceTo(point))
            .FirstOrDefault();

    public static (Connector Source, Connector Target)? GetClosestPair(Element source, Element target)
    {
        var sourceConnectors = ToList(source).Where(IsUsable).ToList();
        var targetConnectors = ToList(target).Where(IsUsable).ToList();
        var best = sourceConnectors
            .SelectMany(s => targetConnectors.Select(t => new { s, t, d = s.Origin.DistanceTo(t.Origin) }))
            .OrderBy(x => x.d)
            .FirstOrDefault();

        return best is null ? null : (best.s, best.t);
    }

    public static IEnumerable<Connector> GetConnectedNeighbours(Element element)
        => ToList(element)
            .Where(c => c.IsConnected)
            .SelectMany(c => c.AllRefs.Cast<Connector>())
            .Where(c => c.Owner.Id != element.Id)
            .GroupBy(c => c.Owner.Id.GetValue()).Select(g => g.First());

    /// <summary>Disconnect all connectors on the given element from neighbours.</summary>
    public static bool TryDisconnectAll(Element element)
    {
        var changed = false;
        foreach (var connector in ToList(element))
        {
            if (!connector.IsConnected) continue;

            foreach (Connector other in connector.AllRefs)
            {
                if (other.Owner.Id == element.Id) continue;
                if (connector.IsConnectedTo(other))
                {
                    connector.DisconnectFrom(other);
                    changed = true;
                }
            }
        }
        return changed;
    }

    /// <summary>
    /// Walk from a seed element along connected MEP graph, collecting all traversed elements.
    /// Useful for "Delete System" style commands.
    /// </summary>
    public static IList<Element> TraverseNetwork(Element seed)
    {
        var visited = new HashSet<long>();
        var result = new List<Element>();
        var queue = new Queue<Element>();
        queue.Enqueue(seed);
        visited.Add(seed.Id.GetValue());

        while (queue.Count > 0)
        {
            var current = queue.Dequeue();
            result.Add(current);

            foreach (var neighbour in GetConnectedNeighbours(current))
            {
                if (visited.Add(neighbour.Owner.Id.GetValue()))
                {
                    queue.Enqueue(neighbour.Owner);
                }
            }
        }
        return result;
    }

    /// <summary>Get the MEP domain of an element (Piping, HVAC, Electrical).</summary>
    public static Domain? GetDomain(Element element)
        => ToList(element).FirstOrDefault(IsUsable)?.Domain;
}

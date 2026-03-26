using System.Collections.Generic;

namespace QuickMechanicalQTO_2.Helpers
{
    public class DuctingQTOResult
    {
        public double TotalInsulationLength { get; set; }  // Tính theo mét (m)

        public List<string> FittingSizes { get; set; } = new List<string>();

        public List<string> AccessoriesSizes { get; set; } = new List<string>();
    }
}

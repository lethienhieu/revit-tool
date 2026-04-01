using System.Collections.Generic;

namespace QuickMechanicalQTO_2.Helpers
{
    public class PipingQTOResult
    {
        /// <summary>
        /// Tổng chiều dài insulation (m) – dùng làm cơ sở thống kê vật tư phụ
        /// </summary>
        public double TotalInsulationLength { get; set; }

        /// <summary>
        /// Danh sách Size của các fittings (VD: DN50, DN100)
        /// </summary>
        public List<string> FittingSizes { get; set; }

        /// <summary>
        /// Danh sách Size của accessories (VD: DN50, DN100)
        /// </summary>
        public List<string> AccessoriesSizes { get; set; }
    }
}

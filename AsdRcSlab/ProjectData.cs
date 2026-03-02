using System;

namespace AsdRcSlab
{
    public class ProjectData
    {
        public string ProjectName { get; set; } = "";
        public string ClientName  { get; set; } = "";
        public string DRWNumber   { get; set; } = "";
        public string Revision    { get; set; } = "P01";
        public int    SlabThickness { get; set; } = 300;
        public double FCK         { get; set; } = 28;
        public double KSpring     { get; set; } = 212333;
        public string ProjectDate { get; set; } = DateTime.Today.ToString("yyyy-MM-dd");

        public string HystoolCode
        {
            get
            {
                switch (SlabThickness)
                {
                    case 225:  return "DK90";
                    case 300:  return "DK165";
                    case 375:  return "DK165";
                    case 450:  return "DK225";
                    case 525:  return "DK225";
                    case 600:  return "DK300";
                    case 675:  return "DK300";
                    case 750:  return "DK375";
                    case 825:  return "DK375";
                    case 900:  return "DK450";
                    case 975:  return "DK450";
                    case 1050: return "DK525";
                    case 1125: return "DK525";
                    case 1200: return "DK600";
                    default:   return "—";
                }
            }
        }
    }
}

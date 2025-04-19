using OfficeOpenXml;

namespace exceltabl
{
    internal class EPPlusLicenseContext : EPPlusLicense
    {
        private object nonCommercial;

        public EPPlusLicenseContext(object nonCommercial)
        {
            this.nonCommercial = nonCommercial;
        }
    }
}
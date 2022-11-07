using System;
using Template4432.Enums;
using Template4432.Models.Base;

namespace Template4432.Models
{
    public class SkiService : Entity
    {
        public string ServiceName { get; set; }
        public string ServiceCode { get; set; }
        
        public SkiServiceType ServiceType { get; set; }
        public decimal PriceForHour { get; set; }
        
        public SkiService() { }

        public SkiService(int id, string serviceName, string serviceCode, string serviceType, decimal priceForHour)
        {
            SkiServiceType? type = serviceType.ToSkiServiceType();

            if (type is null)
                throw new ArgumentException("Нет такого вида услуг", nameof(serviceType));

            Id = id;
            ServiceName = serviceName;
            ServiceCode = serviceCode;
            ServiceType = type.Value;
            PriceForHour = priceForHour;
        }
    }
}
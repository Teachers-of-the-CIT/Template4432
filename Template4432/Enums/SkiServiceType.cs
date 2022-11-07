using System.Runtime.Serialization;
using Newtonsoft.Json;

namespace Template4432.Enums
{
    public enum SkiServiceType
    {
        [EnumMember(Value = "Прокат")]
        Rent,
        
        [EnumMember(Value = "Подъем")]
        Uphill,
        
        [EnumMember(Value = "Обучение")]
        Training
    }

    public static class SkiServiceTypeExtensions
    {
        public static SkiServiceType? ToSkiServiceType(this string str)
        {
            switch (str.Trim())
            {
                case "Прокат":
                {
                    return SkiServiceType.Rent;
                }
                case "Обучение":
                {
                    return SkiServiceType.Training;
                }
                case "Подъем":
                {
                    return SkiServiceType.Uphill;
                }
                default:
                {
                    return null;
                }
            }
        }
    }
}
namespace Template4432.Enums
{
    public enum SkiServiceType
    {
        Rent,
        Uphill,
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
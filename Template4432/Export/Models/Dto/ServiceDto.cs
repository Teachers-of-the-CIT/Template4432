using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Template4432.Export.Models.Dto
{
    [JsonObject]
    public class ServiceDto
    {
        [JsonProperty]
        public int IdServices { get; set; }

        [JsonProperty]
        public string NameServices { get; set; } = string.Empty;

        [JsonProperty]
        public string TypeOfService { get; set; } = string.Empty;

        [JsonProperty]
        public string CodeService { get; set; } = string.Empty;

        [JsonProperty]
        public decimal Cost { get; set; }
    }
}

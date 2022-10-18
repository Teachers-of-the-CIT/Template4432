using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace Template4432.Models
{
    public class JSONClient
    {
        [JsonPropertyName("FullName")]
        public string FIO { get; set; }
        [JsonPropertyName("CodeClient")]
        public string UserId { get; set; }
        public string BirthDate { get; set; }
        public string Index { get; set; }
        public string City { get; set; }
        public string Street { get; set;}
        [JsonPropertyName("Home")]
        public int House { get; set; }
        [JsonPropertyName("Kvartira")]
        public int Apartment { get; set; }
        [JsonPropertyName("E_mail")]
        public string Email { get; set; }
    }
}

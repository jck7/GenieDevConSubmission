using Newtonsoft.Json;

namespace ExcelGenie.Models
{
    public class ErrorResponse
    {
        [JsonProperty("message")]
        public string? Message { get; set; }

        [JsonProperty("validation")]
        public object? Validation { get; set; }

        [JsonProperty("error")]
        public string? Error { get; set; }
    }
} 
using Newtonsoft.Json;

namespace AutoLoanEligibility
{
    public partial class EligibilityModel
    {
        [JsonProperty("layerType")]
        public string LayerType { get; set; }

        [JsonProperty("matchCode")]
        public string MatchCode { get; set; }

        [JsonProperty("eligibilityResult")]
        public string EligibilityResult { get; set; }

        [JsonProperty("entityType")]
        public string EntityType { get; set; }

        [JsonProperty("confidence")]
        public string Confidence { get; set; }

        [JsonProperty("latitude")]
        public string Latitude { get; set; }

        [JsonProperty("postalCode")]
        public string PostalCode { get; set; }

        [JsonProperty("countryRegion")]
        public string CountryRegion { get; set; }

        [JsonProperty("locality")]
        public string Locality { get; set; }

        [JsonProperty("addressLine")]
        public string AddressLine { get; set; }

        [JsonProperty("district")]
        public string District { get; set; }

        [JsonProperty("adminDistrict")]
        public string AdminDistrict { get; set; }

        [JsonProperty("postalTown")]
        public string PostalTown { get; set; }

        [JsonProperty("matchMethod")]
        public string MatchMethod { get; set; }

        [JsonProperty("longitude")]
        public string Longitude { get; set; }
    }
}

using Newtonsoft.Json;

namespace ExportToExcel.Models
{
    public class FeatureResponse
    {
        [JsonProperty("features")]
        public List<TreeFeature> Features { get; set; }
    }

    public class TreeFeature
    {
        [JsonProperty("attributes")]
        public Attributes Attributes { get; set; }
        [JsonProperty("geometry")]
        public Geometry Geometry { get; set; }
    }

    public class Attributes
    {
        [JsonProperty("REQUEST_NO")]
        public long REQUEST_NO{ get; set; }
        [JsonProperty("REQUEST_TYPE")]
        public string REQUEST_TYPE { get; set; }
        [JsonProperty("STATUS_DESCRIPTION")]
        public string STATUS_DESCRIPTION { get; set; }
        
    }

    public class Geometry
    {
        [JsonProperty("x")]
        public double? X { get; set; }
        [JsonProperty("y")]
        public double? Y { get; set; }
    }
}

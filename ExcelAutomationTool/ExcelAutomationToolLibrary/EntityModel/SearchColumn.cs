using Newtonsoft.Json;
using System;

namespace CreateUserReadinessReport.EntityModel
{
    public partial class SearchColumn
    {
        [JsonProperty("New Application ID")]
        public string NewApplicationID { get; set; }
        public string Category { get; set; }
    }
}

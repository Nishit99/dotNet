using Newtonsoft.Json;
using System;

namespace CreateUserReadinessReport.EntityModel
{
    public partial class NewColumn
    {
        [JsonProperty("New Readiness")]
        public string NewReadiness { get; set; }
    }
}

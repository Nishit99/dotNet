using System;
using Newtonsoft.Json;

namespace CreateUserReadinessReport.EntityModel
{
    public partial class BaselineCoumn
    {
        [JsonProperty("ADM ID")]
        public string ADMID { get; set; }

        [JsonProperty("Final Status")]
        public string FinalStatus { get; set; }
    }
}

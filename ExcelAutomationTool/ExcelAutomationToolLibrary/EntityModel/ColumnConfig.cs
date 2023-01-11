using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreateUserReadinessReport.EntityModel
{
    public partial class ColumnConfig
    {

        //public class Rootobject
        //{
        //    public Searchcolumn SearchColumn { get; set; }
        //    public Dynamiccolumn DynamicColumn { get; set; }
        //}

        //public class Searchcolumn
        //{
        //    public string NewApplicationID { get; set; }
        //    public string Category { get; set; }
        //}

        //public class Dynamiccolumn
        //{
        //    public Baselinecoumn BaselineCoumn { get; set; }
        //    public Newcolumn NewColumn { get; set; }
        //}

        //public class Baselinecoumn
        //{
        //    public string ADMID { get; set; }
        //    public string FinalStatus { get; set; }
        //}

        //public class Newcolumn
        //{
        //    public string NewReadiness { get; set; }
        //}

        //public ColumnConfig()
        //{
        //    //SearchColumn = new SearchColumn();
        //    //DynamicColumn = new DynamicColumn();
        //}

        //public SearchColumn SearchColumn { get=>new SearchColumn(); }
        //public DynamicColumn DynamicColumn { get=> new DynamicColumn(); }

        public SearchColumn SearchColumn { get; set; }
        public DynamicColumn DynamicColumn { get; set; }

        
    }

    public partial class ColumnConfig
    {
        public static ColumnConfig FromJson(string json) => JsonConvert.DeserializeObject<ColumnConfig>(json);
    }
}

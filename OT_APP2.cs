//------------------------------------------------------------------------------
// <auto-generated>
//    此代码是根据模板生成的。
//
//    手动更改此文件可能会导致应用程序中发生异常行为。
//    如果重新生成代码，则将覆盖对此文件的手动更改。
// </auto-generated>
//------------------------------------------------------------------------------

namespace Importer
{
    using System;
    using System.Collections.Generic;
    
    public partial class OT_APP2
    {
        public OT_APP2()
        {
            this.OT_WORK = new HashSet<OT_WORK>();
        }
    
        public string APP2_ID { get; set; }
        public string APP2_NO { get; set; }
        public string CYCLE_ID { get; set; }
        public string ORG_ID { get; set; }
        public string STATUS { get; set; }
        public string STATUS_FROM { get; set; }
        public string APP_TYPE { get; set; }
        public string CREATE_BY { get; set; }
        public Nullable<System.DateTime> CREATE_DATE { get; set; }
        public string UPDATE_BY { get; set; }
        public Nullable<System.DateTime> UPDATE_DATE { get; set; }
        public int VERSION_NUM { get; set; }
        public string WORKFLOW_ID { get; set; }
    
        public virtual OT_ORG OT_ORG { get; set; }
        public virtual ICollection<OT_WORK> OT_WORK { get; set; }
    }
}
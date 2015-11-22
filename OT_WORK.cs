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
    
    public partial class OT_WORK
    {
        public OT_WORK()
        {
            this.OT_WORK_OFFSET = new HashSet<OT_WORK_OFFSET>();
        }
    
        public string WORK_ID { get; set; }
        public string WORK_NO { get; set; }
        public string EMP_ID { get; set; }
        public string ORG_ID { get; set; }
        public Nullable<System.DateTime> START_DATE { get; set; }
        public Nullable<System.DateTime> END_DATE { get; set; }
        public Nullable<System.DateTime> CHECK_ON_DATE { get; set; }
        public Nullable<System.DateTime> CHECK_OFF_DATE { get; set; }
        public decimal HOURS { get; set; }
        public string SHIFT_ID { get; set; }
        public decimal TIMES { get; set; }
        public string PAY_TYPE { get; set; }
        public string OT_TYPE { get; set; }
        public string REASON { get; set; }
        public string REMARK { get; set; }
        public string CREATE_BY { get; set; }
        public Nullable<System.DateTime> CREATE_DATE { get; set; }
        public string UPDATE_BY { get; set; }
        public Nullable<System.DateTime> UPDATE_DATE { get; set; }
        public int VERSION_NUM { get; set; }
        public string APP1_ID { get; set; }
        public string APP2_ID { get; set; }
        public Nullable<decimal> HOURS_UPDATE { get; set; }
        public string CREATE_TYPE { get; set; }
        public string REF_WORK_ID { get; set; }
        public string DEL_STATUS { get; set; }
        public Nullable<System.DateTime> ACHIVE_DATE { get; set; }
    
        public virtual OT_APP1 OT_APP1 { get; set; }
        public virtual OT_APP2 OT_APP2 { get; set; }
        public virtual OT_EMP OT_EMP { get; set; }
        public virtual OT_ORG OT_ORG { get; set; }
        public virtual OT_SHIFT OT_SHIFT { get; set; }
        public virtual OT_TYPE OT_TYPE1 { get; set; }
        public virtual ICollection<OT_WORK_OFFSET> OT_WORK_OFFSET { get; set; }
    }
}

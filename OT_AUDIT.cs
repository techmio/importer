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
    
    public partial class OT_AUDIT
    {
        public string AUDIT_ID { get; set; }
        public int SEQ_ID { get; set; }
        public string AUDITER { get; set; }
        public string REMARKS { get; set; }
        public string STATUS { get; set; }
        public string APP_ID { get; set; }
        public string APP_TYPE { get; set; }
        public string CREATE_BY { get; set; }
        public Nullable<System.DateTime> CREATE_DATE { get; set; }
        public string UPDATE_BY { get; set; }
        public Nullable<System.DateTime> UPDATE_DATE { get; set; }
        public int VERSION_NUM { get; set; }
    
        public virtual OT_EMP OT_EMP { get; set; }
    }
}
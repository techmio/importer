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
    
    public partial class OT_WORK_OFFSET
    {
        public string WORK_OFFSET { get; set; }
        public string WORK_ID { get; set; }
        public string OFFSET_ID { get; set; }
        public Nullable<decimal> OFFSET_HOURS { get; set; }
        public Nullable<decimal> WORK_HOURS { get; set; }
    
        public virtual OT_OFFSET OT_OFFSET { get; set; }
        public virtual OT_WORK OT_WORK { get; set; }
    }
}

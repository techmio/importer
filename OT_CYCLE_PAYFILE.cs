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
    
    public partial class OT_CYCLE_PAYFILE
    {
        public string CYCLE_PAYFILE_ID { get; set; }
        public string CYCLE_ID { get; set; }
        public byte[] OT_FILE { get; set; }
        public System.DateTime CREATE_DATE { get; set; }
        public System.DateTime UPLOAD_DATE { get; set; }
    
        public virtual OT_CYCLE OT_CYCLE { get; set; }
    }
}
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Importer
{
       public abstract class ISysSeqService {
	    public   const String SEQ_CODE_APP1_PERSON = "APP1_PERSON";
        public  const String SEQ_CODE_APP1_COLLECT = "APP1_COLLECT";
        public  const String SEQ_CODE_APP2 = "APP2";
        public  const String SEQ_CODE_APP3 = "APP3";
        public  const String SEQ_CODE_APP4 = "APP4";
        public  const String SEQ_CODE_WORK = "OT_WORK";
        public  const String SEQ_CODE_OFFSET = "OT_OFFSET";
	    public abstract long findSeqNumByCode(String code);
	    public abstract String getSeqNumByCode(String code);
    }
}

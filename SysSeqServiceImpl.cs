using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Importer
{
    class SysSeqServiceImpl:ISysSeqService
    {

        public override long  findSeqNumByCode(string code)
        {
            OTEntities context = new OTEntities();
            var q = from p in context.SYS_SEQ where p.SEQ_CODE == code select p;
            if(q.FirstOrDefault()!=null)
            {
                SYS_SEQ seq = q.FirstOrDefault();
                long num =(long)seq.SEQ_NUM;
                seq.SEQ_NUM ++;
                context.SaveChanges();
                return num;
            }
           /* SysSeq ss = sysSeqDao.findBySeqCode(code);
            long seqNum = ss.getSeqNum();
            ss.setSeqNum(ss.getSeqNum() + 1);
            sysSeqDao.save(ss);
            return seqNum;*/
            return 0;
        }

        

        public string getNumberByCode(string code,DateTime dt)
        {
            long seqNum = this.findSeqNumByCode(code);
            string prefix = "";
            string year = dt.Year.ToString().Substring(2, 2);
            switch (code)
            {
                //申请单号
                case ISysSeqService.SEQ_CODE_APP1_PERSON:
                    prefix += "OB";
                    break;
                case ISysSeqService.SEQ_CODE_APP1_COLLECT:
                    prefix += "OBG";
                    break;
                case ISysSeqService.SEQ_CODE_APP2:
                    prefix += "A2";
                    break;
                //
                case ISysSeqService.SEQ_CODE_APP3:
                    prefix += "LB";
                    break;
                case ISysSeqService.SEQ_CODE_APP4:
                    prefix += "A4";
                    break;
                //加班号
                case ISysSeqService.SEQ_CODE_WORK:
                    prefix += "OD";
                    break;
                //offset number    
                case ISysSeqService.SEQ_CODE_OFFSET:
                    prefix += "LD";
                    break;
                default:
                    break;
            }
            prefix += year;
            int month = dt.Month;
            String monthStr = month + "";
            if (month == 10)
            {
                monthStr = "A";
            }
            else if (month == 11)
            {
                monthStr = "B";
            }
            else if (month == 12)
            {
                monthStr = "C";
            }
            prefix += monthStr;


            String seqN = seqNum.ToString().PadLeft(6, '0');
            return prefix + seqN;

        }

        public override string getSeqNumByCode(string code)
        {
            long seqNum = this.findSeqNumByCode(code);
            string prefix = "";
            string year = DateTime.Now.Year.ToString().Substring(2, 2);
            switch(code)
            {
                //申请单号
                case ISysSeqService.SEQ_CODE_APP1_PERSON:
                    prefix += "OB";
                    break;
                case ISysSeqService.SEQ_CODE_APP1_COLLECT:
                    prefix += "OBG";
                    break;
                case ISysSeqService.SEQ_CODE_APP2:
                    prefix += "A2";
                    break;
                //
                case ISysSeqService.SEQ_CODE_APP3:
                    prefix += "LB";
                    break;
                case ISysSeqService.SEQ_CODE_APP4:
                    prefix+="A4";
                    break;
                //加班号
                case ISysSeqService.SEQ_CODE_WORK:
                    prefix += "OD";
                    break;
                //offset number    
                case ISysSeqService.SEQ_CODE_OFFSET:
                    prefix += "LD";
                    break;
                default:
                    break;
            }
            prefix += year;
            int month = DateTime.Now.Month;
            String monthStr = month + "";
            if (month == 10)
            {
                monthStr = "A";
            }
            else if (month == 11)
            {
                monthStr = "B";
            }
            else if (month == 12)
            {
                monthStr = "C";
            }
            prefix += monthStr;


            String seqN = seqNum.ToString().PadLeft(6,'0');
            return prefix + seqN;

           


            /*
             Long seqNum = this.findSeqNumByCode(code);
		        String prifix = "";
		        Calendar c = Calendar.getInstance();
		        String year = c.get(Calendar.YEAR) + "";
		        year = year.substring(year.length() - 2, year.length());
		        if(ISysSeqService.SEQ_CODE_APP1_PERSON.equals(code)){
			        prifix += "OB";
		        }else if(ISysSeqService.SEQ_CODE_APP1_COLLECT.equals(code)){
			        prifix += "OBG";
		        }else if(ISysSeqService.SEQ_CODE_APP2.equals(code)){
			        prifix += "A2";
		        }else if(ISysSeqService.SEQ_CODE_APP3.equals(code)){
			        prifix += "LB";
		        }else if(ISysSeqService.SEQ_CODE_APP4.equals(code)){
			        prifix +="A4";
		        }else if(ISysSeqService.SEQ_CODE_WORK.equals(code)){
			        prifix += "OD";
		        }else if(ISysSeqService.SEQ_CODE_OFFSET.equals(code)){
			        prifix += "LD";
		        }
		        prifix += year;
		        int month = c.get(Calendar.MONTH)+1;
		        String monthStr = month+"";
		        if(month == 10){
			        monthStr = "A";
		        }else if(month == 11){
			        monthStr = "B";
		        }else if(month == 12){
			        monthStr = "C";
		        }
		        prifix += monthStr;
		        String seqN = StringUtils.leftPad(seqNum+"", 6, '0');
		        return prifix + seqN;
             */
            
        }
    }
}

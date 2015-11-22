using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;
using Oracle.DataAccess.Client;
using System.Collections;

namespace Importer
{
   public class OTControler
    {
        public OTEntities context = new OTEntities();
        public string GroupId = "";
        public string GroupName = "";
        public string DeptId = "";
        public string DeptName = "";
        public String EmpId = "";
        public const string DATEFORMAT = "yyyy/M/d h:mm tt";
        public const string ONLYDATEFORMAT = "yyyy/M/d h:mm:ss";
        public const string NOTIMEFORMAT = "yyyy/m/d";
        SysSeqServiceImpl SYS = new SysSeqServiceImpl();
        CultureInfo ZH = new CultureInfo("zh-CN");

        public OTControler()
        {
           
            context.SYS_SEQ.AsNoTracking();
            context.OT_WORK.AsNoTracking();
            context.OT_SHIFT.AsNoTracking();
            context.OT_CYCLE.AsNoTracking();
            context.OT_APP1.AsNoTracking();
            context.OT_APP2.AsNoTracking();
            context.OT_APP3.AsNoTracking();
            context.OT_OFFSET.AsNoTracking();
            context.OT_WORK_OFFSET.AsNoTracking();
            

        }

        public static List<OT_SHIFT> listShift = new List<OT_SHIFT>();
        public static List<OT_APP1> listApp1 = new List<OT_APP1>();
        public static List<OT_APP2> listApp2 = new List<OT_APP2>();
        public static List<OT_WORK> listOTWork = new List<OT_WORK>();

        public static Hashtable EmpAbnormal = new Hashtable();

        public bool CheckEmployeeNumber(string employeeName, string employeeNumber, string groupName, string deptName)
        {
            if (Comm.NoExistEmp.Keys.Contains(employeeNumber)) return false;
            var q = from p in context.OT_EMP where p.EMP_NUMBER == employeeNumber select p;
            OT_EMP employee = q.FirstOrDefault();
            if (employee == null)
            {
                if (!Comm.NoExistEmp.Keys.Contains(employeeNumber))
                    Comm.NoExistEmp.Add(employeeNumber, string.Format("库不存在员工{0},员工号码{1},所在的组为-{2},部门-{3}", employeeName, employeeNumber, groupName, deptName));
                return false;
            }
            return true;
        }

        
        public bool CheckEmployee(string employeeName, string employeeNumber, string groupName, string deptName)
        {
            ProblemEmployee pe = new ProblemEmployee();
            pe.Name = employeeName;
            pe.Number = employeeNumber;
            pe.OldDeptName = deptName;
            pe.OldGroupName = groupName;
            
         
            int deptNumber = 0, groupNumber = 0;
            ExcelTool.GetExcelDeptGroupID(deptName, groupName, ref deptNumber, ref groupNumber, ref this.DeptName, ref this.GroupName);
            pe.OldDeptId = deptNumber.ToString();
            pe.OldGroupId = groupNumber.ToString();
            pe.DeptName = this.DeptName;
            pe.GroupName = this.GroupName;

            OT_POSITION position = null; ;
            bool matchMappingRecord = true;
            string msg = String.Format("Excel中员工{0},员工号码{1},于{2},{3}", employeeName, employeeNumber, deptName, groupName); ;
            if (groupNumber>0)
            {
                var q1 = from p in context.OT_POSITION where p.OT_EMP.EMP_NUMBER == employeeNumber && p.OT_ORG.ORG_NUM == groupNumber select p;
                position = q1.FirstOrDefault();
                if (position != null)
                {
                    pe.GroupId = position.OT_ORG.ORG_NUM.ToString(); ;
                    pe.GroupName = position.OT_ORG.ORG_NAME;
                    
                }
                //else
                //    Comm.Logger.Info("员工信息与文件一致");
            }

            if (position == null)
            {
                matchMappingRecord = false;
                var q1 = from p in context.OT_POSITION where p.OT_EMP.EMP_NUMBER == employeeNumber select p;
                position = q1.FirstOrDefault();
                if (position == null)
                    return false;
            }

            this.GroupId = position.OT_ORG.ORG_ID.ToString();
            pe.GroupId = position.OT_ORG.ORG_NUM.ToString();
            pe.GroupName = position.OT_ORG.ORG_NAME;
            this.GroupName = position.OT_ORG.ORG_NAME.ToString();
           
            this.EmpId = position.EMP_ID;
            var q2 = from p in context.OT_ORG where p.ORG_ID == position.OT_ORG.PARANT_ORG select p;
            OT_ORG department = q2.FirstOrDefault();
            if (department == null)
            {
                Comm.Logger.Error(string.Format("该员工没有分配部门{0},员工号码{1}", employeeName, employeeNumber));
                return false;
             }
            this.DeptId = department.ORG_ID;
            pe.DeptName=this.DeptName = department.ORG_NAME;
            pe.DeptId = department.ORG_NUM.ToString();
            
            if ((!department.ORG_NUM.Equals(deptNumber) && deptNumber != 0)
                ||(!pe.OldGroupId.Equals(pe.GroupId)))
            {
                msg = string.Format("Excel文件员工{0},员工号码{1},部门{2}-{3}不同其在库中的部门{4}-{5}", employeeName, employeeNumber,
                    deptNumber, deptName, department.ORG_NUM, department.ORG_NAME);
                Comm.Logger.Warn(msg);
                if (!EmpAbnormal.ContainsKey(employeeNumber))
                {
                    EmpAbnormal.Add(employeeNumber, employeeName);
                    ExcelTool.WriteErrorOLEDB(pe);
                }
            }


           

            return true;

        }


        public void CheckHaveTwoORG()
        {
            foreach (OT_POSITION emp in context.OT_POSITION)
            {
                var q = from p in context.OT_POSITION  where p.EMP_ID == emp.EMP_ID select p;
                if(q.Count()>1)
                {
                    Comm.Logger.Info(string.Format("EMP_ID {0} in  {1}  orgs ", emp.EMP_ID, q.Count()));
                }
            }

        }


        public void CheckOffsetIntegrated()
        {
          
               
                var q = from p in context.OT_WORK where p.PAY_TYPE=="1" select p;
                Comm.Logger.Info(string.Format("等待确认记录共{0}条.", q.Count()));
                foreach (OT_WORK otwork in q)
                {
                    decimal offset_hours = 0;
                   foreach(OT_WORK_OFFSET work_offset in otwork.OT_WORK_OFFSET)
                   {
                       offset_hours +=(decimal)work_offset.OFFSET_HOURS;
                       
                   }
                   if(!otwork.HOURS.Equals(offset_hours))
                   {
                       ProblemOT pot = new ProblemOT();
                       pot.EmpName = otwork.OT_EMP.EMP_NAME;
                       pot.EmpNumber = otwork.OT_EMP.EMP_NUMBER;
                       pot.Offset_hours = ((DateTime)otwork.START_DATE).ToString("yyyy-MM-dd");
                       pot.Apply_Number = otwork.OT_APP1.APP1_NO;
                       pot.OT_Hours = otwork.HOURS.ToString();
                       pot.Offset_hours = offset_hours.ToString();                       
                       ExcelTool.WriteErrorOLEDB(pot);
                   }
                }

        }

       
        public bool CheckEmployeeExist(OTItem item)
        {

            if (!CheckEmployeeNumber(item.Worker_CnName, item.Worker_Number, item.Worker_Group, item.Worker_Dept))
                return false;
            return true;
        }

        public bool CheckEmployeeExist(OffsetItem item)
        {

            if (!CheckEmployeeNumber(item.Worker_CnName, item.Worker_Number, item.Worker_Group, item.Worker_Dept))
                return false;
            return true;
        }







       public void SaveExcelOTs(OTItem item)
       {
           try
           {
               /*
               OT_WORK_EXCEL rec=new OT_WORK_EXCEL();
               rec.ID = Guid.NewGuid().ToString();
               rec.emp_number = item.Worker_Number;
               rec.emp_name = item.Worker_CnName;
               rec.emp_group = item.Worker_Group;
               rec.emp_dept = item.Worker_Dept;

               DateTime started = DateTime.ParseExact(item.Cycle_StartEd, ONLYDATEFORMAT, ZH);
               DateTime ended = DateTime.ParseExact(item.Cycle_EndEd, ONLYDATEFORMAT, ZH);

               rec.cycle_started = started;
               rec.cycle_ended = ended;

               rec.ot_ed = DateTime.ParseExact(item.OT_StartEd, ONLYDATEFORMAT, ZH);

               string strDate=item.OT_StartEd;
               if (item.OT_StartTime.Length < DATEFORMAT.Length)
                 strDate = item.OT_StartEd.ToString().Substring(0,NOTIMEFORMAT.Length+1).Trim() + " " + item.OT_StartTime;
               rec.ot_start = DateTime.ParseExact(strDate, DATEFORMAT,ZH);

               strDate = item.OT_EndTime;
               if (strDate.Length < DATEFORMAT.Length)
                   strDate = item.OT_EndEd.Substring(0, NOTIMEFORMAT.Length + 1).Trim() + " " + item.OT_EndTime;
               rec.ot_end= DateTime.ParseExact(strDate, DATEFORMAT, ZH);

               rec.work_hour = decimal.Parse(item.OT_Hours);
               rec.pay_hour = decimal.Parse(item.Pay_Hours);
               rec.offset_hour = decimal.Parse(item.Offset_Hours);
               rec.reason = item.Reason;
               rec.shift_name = item.Shift_Id;
               rec.compensate_rate = decimal.Parse(item.Compensate_Rate);

               context.OT_WORK_EXCEL.Add(rec);
               context.SaveChanges();
                * */
              
               
           }
           catch (Exception e)
           {
               Comm.Logger.Error(e.Message);
           }
       }


      public bool ProcessOTItemNoDB(OTItem item)
       {
           OT_WORK otwork = new OT_WORK();
           otwork.WORK_ID = Guid.NewGuid().ToString();
           otwork.VERSION_NUM = 0;
           otwork.HOURS_UPDATE = 0;
           otwork.CREATE_DATE = DateTime.Parse(item.Create_Ed);
           otwork.CREATE_BY = "System";
           SysSeqServiceImpl SYS = new SysSeqServiceImpl();


           if (string.IsNullOrEmpty(item.OT_ApplyNumber)||item.OT_ApplyNumber=="")
           {
               //旧系统没有此单号则新系统导入时根据申请的批次或申报周期自动生成
               OT_APP1 app1 = CreateAPP1NoDB(item);
               otwork.APP1_ID = app1.APP1_ID;
               OT_APP2 app2 = CreateAPP2NoDB(item, app1.CYCLE_ID);
               otwork.APP2_ID = app2.APP2_ID;

           }
           else
           {
               var q = from p in context.OT_APP1 where p.APP1_NO == item.OT_ApplyNumber select p;
               OT_APP1 existApp1 = q.FirstOrDefault();
               if (existApp1 == null)
               {
                   Comm.Logger.Warn(string.Format("No exist app1_no in database with value {0}", item.OT_ApplyNumber));
                   OT_APP1 app1 = CreateAPP1NoDB(item);
                   otwork.APP1_ID = app1.APP1_ID;
                   OT_APP2 app2 = CreateAPP2NoDB(item, app1.CYCLE_ID);
                   otwork.APP2_ID = app2.APP2_ID;

               }
               else
               {
                   otwork.APP1_ID = existApp1.APP1_ID;
               }
           }
           DateTime dt_app1app2 = DateTime.Now;
           //一二级审批人, october vesrtion no need
           // CreateAudit(item, otwork.APP1_ID);
           if (!string.IsNullOrEmpty(item.OT_Number))
               otwork.WORK_NO = item.OT_Number;
           else
               otwork.WORK_NO = SYS.getNumberByCode(ISysSeqService.SEQ_CODE_WORK, (DateTime)otwork.CREATE_DATE);

           otwork.ACHIVE_DATE = DateTime.Now; //store current date as achive date
           otwork.DEL_STATUS = "0";

           otwork.HOURS = decimal.Parse(item.OT_Hours);
           otwork.TIMES = decimal.Parse(item.Compensate_Rate);

           string strDate = item.OT_StartTime;
           if (item.OT_StartTime.Length < DATEFORMAT.Length)
               strDate = item.OT_StartEd.ToString().Substring(0, NOTIMEFORMAT.Length + 1).Trim() + " " + item.OT_StartTime;
           otwork.START_DATE = DateTime.ParseExact(strDate, DATEFORMAT, ZH);
           strDate = item.OT_EndTime;
           if (strDate.Length < DATEFORMAT.Length)
               strDate = item.OT_EndEd.Substring(0, NOTIMEFORMAT.Length + 1).Trim() + " " + item.OT_EndTime;
           otwork.END_DATE = DateTime.ParseExact(strDate, DATEFORMAT, ZH);
           otwork.OT_TYPE = "-1"; //-1=从旧系统导入
          
          otwork.PAY_TYPE = "0"; // 0 or 1?
          decimal pay_hours=string.IsNullOrEmpty(item.Pay_Hours)?0:item.Pay_Hours==""?0:decimal.Parse(item.Pay_Hours);
          decimal offset_hours = string.IsNullOrEmpty(item.Offset_Hours) ? 0 : item.Offset_Hours == "" ? 0 : decimal.Parse(item.Offset_Hours);
          if (!offset_hours.Equals(otwork.HOURS - pay_hours))
          {
              Comm.Logger.Error("补休!=加班-支付");
              ExcelTool.WriteErrorOLEDB(item, "补休!=加班-支付");            
          }

          if(pay_hours==0)
             otwork.PAY_TYPE = "1"; 
         

           var q2 = from p in context.OT_SHIFT where p.SHIFT_NAME == item.Shift_Id select p;
           if (q2.FirstOrDefault() == null)
           {
               Comm.Logger.Warn(string.Format("没发现班次{0}存在,新建了班次", item.Shift_Id));
               otwork.SHIFT_ID = CreateNewShift(item.Shift_Id, item.Compensate_Rate).SHIFT_ID;
           }
           else
               otwork.SHIFT_ID = q2.FirstOrDefault().SHIFT_ID;
           otwork.ORG_ID = this.GroupId;
           otwork.EMP_ID = this.EmpId;
           otwork.VERSION_NUM = 0;
           otwork.CREATE_TYPE = "0";

           otwork.REASON = item.Reason;
           otwork.REMARK = item.Comment + "-从旧系统导入";
           listOTWork.Add(otwork);
           return true;

       }
     

        public bool ProcessOTItem(OTItem item)
        {          

            OT_WORK otwork = new OT_WORK();
            otwork.WORK_ID = Guid.NewGuid().ToString();
            otwork.VERSION_NUM = 0;
            otwork.CREATE_DATE = DateTime.Parse(item.Create_Ed);
            otwork.CREATE_BY = "System";
            SysSeqServiceImpl SYS = new SysSeqServiceImpl();


            if (string.IsNullOrEmpty(item.OT_ApplyNumber))
            {
                //旧系统没有此单号则新系统导入时根据申请的批次或申报周期自动生成
                OT_APP1 app1 = CreateAPP1(item);
                otwork.APP1_ID = app1.APP1_ID;
                OT_APP2 app2 = CreateAPP2(item, app1.CYCLE_ID);
                otwork.APP2_ID = app2.APP2_ID;
                
            }
            else
            {
                var q = from p in context.OT_APP1 where p.APP1_NO == item.OT_ApplyNumber select p;
                if (q.FirstOrDefault()==null)
                {
                    Comm.Logger.Warn(string.Format("No exist app1_no in database with value {0}", item.OT_ApplyNumber));
                    OT_APP1 app1 = CreateAPP1(item);
                    otwork.APP1_ID = app1.APP1_ID;
                    OT_APP2 app2 = CreateAPP2(item, app1.CYCLE_ID);
                    otwork.APP2_ID = app2.APP2_ID;
                }
                else
                {
                    otwork.APP1_ID = q.FirstOrDefault().APP1_ID;
                }
            }
            DateTime dt_app1app2 = DateTime.Now;
            //一二级审批人, october vesrtion no need
           // CreateAudit(item, otwork.APP1_ID);
            if (!string.IsNullOrEmpty(item.OT_Number))
                otwork.WORK_NO = item.OT_Number;
            else
                otwork.WORK_NO = SYS.getNumberByCode(ISysSeqService.SEQ_CODE_WORK,(DateTime)otwork.CREATE_DATE);
           
            otwork.ACHIVE_DATE = DateTime.Now; //store current date as achive date
            otwork.DEL_STATUS = "0";

            otwork.HOURS = decimal.Parse(item.OT_Hours);
            otwork.TIMES = decimal.Parse(item.Compensate_Rate);

            string strDate = item.OT_StartTime;
            if (item.OT_StartTime.Length < DATEFORMAT.Length)
                strDate = item.OT_StartEd.ToString().Substring(0,NOTIMEFORMAT.Length+1).Trim() + " " + item.OT_StartTime;
            otwork.START_DATE = DateTime.ParseExact(strDate, DATEFORMAT,ZH);
            strDate = item.OT_EndTime;
            if (strDate.Length < DATEFORMAT.Length)
                strDate = item.OT_EndEd.Substring(0,NOTIMEFORMAT.Length+1).Trim() + " " + item.OT_EndTime;
            otwork.END_DATE = DateTime.ParseExact(strDate, DATEFORMAT, ZH);
            otwork.OT_TYPE = "-1"; //-1=从旧系统导入
            otwork.PAY_TYPE = "0"; // 0 or 1?
            decimal pay_hours = string.IsNullOrEmpty(item.Pay_Hours) ? 0 : item.Pay_Hours == "" ? 0 : decimal.Parse(item.Pay_Hours);
            decimal offset_hours = string.IsNullOrEmpty(item.Offset_Hours) ? 0 : item.Offset_Hours == "" ? 0 : decimal.Parse(item.Offset_Hours);
            if (!offset_hours.Equals(otwork.HOURS - pay_hours))
            {
                Comm.Logger.Error("补休!=加班-支付");
                ExcelTool.WriteErrorOLEDB(item, "补休!=加班-支付");
               
            }

            if (pay_hours == 0)
                otwork.PAY_TYPE = "1"; 
         

            var q2=from p in context.OT_SHIFT where p.SHIFT_NAME==item.Shift_Id select p;
            if (q2.FirstOrDefault() == null)
            {
                 Comm.Logger.Warn(string.Format("没发现班次{0}存在,新建了班次", item.Shift_Id));
                 otwork.SHIFT_ID =CreateNewShift(item.Shift_Id, item.Compensate_Rate).SHIFT_ID;                
            }
            else
            otwork.SHIFT_ID = q2.FirstOrDefault().SHIFT_ID;
            otwork.ORG_ID = this.GroupId;
            otwork.EMP_ID = this.EmpId;
            otwork.VERSION_NUM = 0;
            otwork.CREATE_TYPE = "0";
            otwork.REASON = item.Reason;
            otwork.REMARK=item.Comment+"-从旧系统导入";
            context.OT_WORK.Add(otwork);            
            Comm.Logger.Info("记录入库.");
            return true;
            
        }


        public static int BaseShiftSequence=-1;
        OT_SHIFT CreateNewShiftNoDB(string shift_name, string compensate_rate)
        {
            if (BaseShiftSequence.Equals(-1))
            {
                var q = from p in context.OT_SHIFT orderby p.SEQ_ID descending select p;
                int seq = 1;
                if (q.FirstOrDefault() != null)
                {
                    seq = (short)q.FirstOrDefault().SEQ_ID+1;
                }
                BaseShiftSequence = seq;
            }
            OT_SHIFT shift = new OT_SHIFT();
            shift.SHIFT_ID = Guid.NewGuid().ToString();
            shift.SHIFT_NAME = shift_name;
            shift.STATUS = "0";
            shift.CREATE_BY = "System";
            shift.CREATE_DATE = DateTime.Now;
            shift.OT_TIME = decimal.Parse(compensate_rate);
            shift.SEQ_ID = (short)BaseShiftSequence;
            shift.VERSION_NUM = 0;
            BaseShiftSequence++;
            return shift;
        }


        OT_SHIFT CreateNewShift(string shift_name,string compensate_rate)
        {
            var q=from p in context.OT_SHIFT orderby p.SEQ_ID descending select p;
            OT_SHIFT shift= new OT_SHIFT();
            shift.SHIFT_ID = Guid.NewGuid().ToString();
            shift.SHIFT_NAME = shift_name;
            shift.STATUS = "0";
            shift.CREATE_BY = "System";
            shift.CREATE_DATE = DateTime.Now;
            shift.OT_TIME = decimal.Parse(compensate_rate);
            int seq=1;
            if (q.FirstOrDefault() != null)             
            {
                if (q.FirstOrDefault().SEQ_ID != null)
                    seq = (int)q.FirstOrDefault().SEQ_ID + 1;
                else
                    seq = 1;
            }
            shift.SEQ_ID =(short) seq ;
            shift.VERSION_NUM = 0;
            context.OT_SHIFT.Add(shift);
            context.SaveChanges();
            return shift;
            
        }

        public void SaveOTWork()
        {
            if (listOTWork.Count() == 0) return;
            OracleConnection conn = new OracleConnection(context.Database.Connection.ConnectionString);
            OracleCommand command = new OracleCommand();
            command.Connection = conn;
            command.ArrayBindCount = listOTWork.Count();
            command.CommandText = @"insert into OT_WORK values(:WORK_ID,              
              :WORK_NO,              
              :EMP_ID,              
              :ORG_ID,              
              :START_DATE,              
              :END_DATE,              
              :CHECK_ON_DATE,            
              :CHECK_OFF_DATE,              
              :HOURS,              
              :SHIFT_ID,
              :TIMES,
              :PAY_TYPE,
              :OT_TYPE,
              :REASON,
              :REMARK,
              :CREATE_BY,
              :CREATE_DATE,
              :UPDATE_BY,
              :UPDATE_DATE,
              :VERSION_NUM,
              :APP1_ID,
              :APP2_ID,
              :HOURS_UPDATE,
              :CREATE_TYPE,
              :REF_WORK_ID,
              :DEL_STATUS,
              :ACHIVE_DATE)";
            conn.Open();
            OracleParameter WORK_ID_Param = new OracleParameter("WORK_ID", OracleDbType.Varchar2);
            WORK_ID_Param.Direction = ParameterDirection.Input;
            WORK_ID_Param.Value = (from p in listOTWork select p.WORK_ID).ToArray();
            command.Parameters.Add(WORK_ID_Param);

            OracleParameter WORK_NO_Param = new OracleParameter("WORK_NO", OracleDbType.Varchar2);
            WORK_NO_Param.Direction = ParameterDirection.Input;
            WORK_NO_Param.Value = (from p in listOTWork select p.WORK_NO).ToArray();
            command.Parameters.Add(WORK_NO_Param);

            OracleParameter EMP_ID_Param = new OracleParameter("EMP_ID", OracleDbType.Varchar2);
            EMP_ID_Param.Direction = ParameterDirection.Input;
            EMP_ID_Param.Value = (from p in listOTWork select p.EMP_ID).ToArray();
            command.Parameters.Add(EMP_ID_Param);

            OracleParameter ORG_ID_Param = new OracleParameter("ORG_ID", OracleDbType.Varchar2);
            ORG_ID_Param.Direction = ParameterDirection.Input;
            ORG_ID_Param.Value = (from p in listOTWork select p.ORG_ID).ToArray();
            command.Parameters.Add(ORG_ID_Param);

            OracleParameter START_DATE_Param = new OracleParameter("START_DATE", OracleDbType.Date);
            START_DATE_Param.Direction = ParameterDirection.Input;
            START_DATE_Param.Value = (from p in listOTWork select p.START_DATE).ToArray();
            command.Parameters.Add(START_DATE_Param);

            OracleParameter END_DATE_Param = new OracleParameter("END_DATE", OracleDbType.Date);
            END_DATE_Param.Direction = ParameterDirection.Input;
            END_DATE_Param.Value = (from p in listOTWork select p.END_DATE).ToArray();
            command.Parameters.Add(END_DATE_Param);

            OracleParameter CHECK_ON_DATE_Param = new OracleParameter("CHECK_ON_DATE", OracleDbType.Date);
            CHECK_ON_DATE_Param.Direction = ParameterDirection.Input;
            CHECK_ON_DATE_Param.Value = (from p in listOTWork select p.CHECK_ON_DATE).ToArray();
            command.Parameters.Add(CHECK_ON_DATE_Param);

            OracleParameter CHECK_OFF_DATE_Param = new OracleParameter("CHECK_OFF_DATE", OracleDbType.Date);
            CHECK_OFF_DATE_Param.Direction = ParameterDirection.Input;
            CHECK_OFF_DATE_Param.Value = (from p in listOTWork select p.CHECK_OFF_DATE).ToArray();
            command.Parameters.Add(CHECK_OFF_DATE_Param);

            OracleParameter HOURS_Param = new OracleParameter("HOURS", OracleDbType.Decimal);
            HOURS_Param.Direction = ParameterDirection.Input;
            HOURS_Param.Value = (from p in listOTWork select p.HOURS).ToArray();
            command.Parameters.Add(HOURS_Param);

            OracleParameter SHIFT_ID_Param = new OracleParameter("SHIFT_ID", OracleDbType.Varchar2);
            SHIFT_ID_Param.Direction = ParameterDirection.Input;
            SHIFT_ID_Param.Value = (from p in listOTWork select p.SHIFT_ID).ToArray();
            command.Parameters.Add(SHIFT_ID_Param);

            OracleParameter TIMES_Param = new OracleParameter("TIMES", OracleDbType.Varchar2);
            TIMES_Param.Direction = ParameterDirection.Input;
            TIMES_Param.Value = (from p in listOTWork select p.TIMES).ToArray();
            command.Parameters.Add(TIMES_Param);

            OracleParameter PAY_TYPE_Param = new OracleParameter("PAY_TYPE", OracleDbType.Varchar2);
            PAY_TYPE_Param.Direction = ParameterDirection.Input;
            PAY_TYPE_Param.Value = (from p in listOTWork select p.PAY_TYPE).ToArray();
            command.Parameters.Add(PAY_TYPE_Param);

            OracleParameter OT_TYPE_Param = new OracleParameter("OT_TYPE", OracleDbType.Varchar2);
            OT_TYPE_Param.Direction = ParameterDirection.Input;
            OT_TYPE_Param.Value = (from p in listOTWork select p.OT_TYPE).ToArray();
            command.Parameters.Add(OT_TYPE_Param);

            OracleParameter REASON_Param = new OracleParameter("REASON", OracleDbType.Varchar2);
            REASON_Param.Direction = ParameterDirection.Input;
            REASON_Param.Value = (from p in listOTWork select p.REASON).ToArray();
            command.Parameters.Add(REASON_Param);

            OracleParameter REMARK_Param = new OracleParameter("REMARK", OracleDbType.Varchar2);
            REMARK_Param.Direction = ParameterDirection.Input;
            REMARK_Param.Value = (from p in listOTWork select p.REMARK).ToArray();
            command.Parameters.Add(REMARK_Param);

            OracleParameter CREATE_BY_Param = new OracleParameter("CREATE_BY", OracleDbType.Varchar2);
            CREATE_BY_Param.Direction = ParameterDirection.Input;
            CREATE_BY_Param.Value = (from p in listOTWork select p.CREATE_BY).ToArray();
            command.Parameters.Add(CREATE_BY_Param);

            OracleParameter CREATE_DATE_Param = new OracleParameter("CREATE_DATE", OracleDbType.Date);
            CREATE_DATE_Param.Direction = ParameterDirection.Input;
            CREATE_DATE_Param.Value = (from p in listOTWork select p.CREATE_DATE).ToArray();
            command.Parameters.Add(CREATE_DATE_Param);

            OracleParameter UPDATE_BY_Param = new OracleParameter("UPDATE_BY", OracleDbType.Varchar2);
            UPDATE_BY_Param.Direction = ParameterDirection.Input;
            UPDATE_BY_Param.Value = (from p in listOTWork select p.UPDATE_BY).ToArray();
            command.Parameters.Add(UPDATE_BY_Param);

            OracleParameter UPDATE_DATE_Param = new OracleParameter("UPDATE_DATE", OracleDbType.Date);
            UPDATE_DATE_Param.Direction = ParameterDirection.Input;
            UPDATE_DATE_Param.Value = (from p in listOTWork select p.UPDATE_DATE).ToArray();
            command.Parameters.Add(UPDATE_DATE_Param);

            OracleParameter VERSION_NUM_Param = new OracleParameter("VERSION_NUM", OracleDbType.Int16);
            VERSION_NUM_Param.Direction = ParameterDirection.Input;
            VERSION_NUM_Param.Value = (from p in listOTWork select p.VERSION_NUM).ToArray();
            command.Parameters.Add(VERSION_NUM_Param);

            OracleParameter APP1_ID_Param = new OracleParameter("APP1_ID", OracleDbType.Varchar2);
            APP1_ID_Param.Direction = ParameterDirection.Input;
            APP1_ID_Param.Value = (from p in listOTWork select p.APP1_ID).ToArray();
            command.Parameters.Add(APP1_ID_Param);

            OracleParameter APP2_ID_Param = new OracleParameter("APP2_ID", OracleDbType.Varchar2);
            APP2_ID_Param.Direction = ParameterDirection.Input;
            APP2_ID_Param.Value = (from p in listOTWork select p.APP2_ID).ToArray();
            command.Parameters.Add(APP2_ID_Param);

            OracleParameter HOURS_UPDATE_Param = new OracleParameter("HOURS_UPDATE", OracleDbType.Decimal);
            HOURS_UPDATE_Param.Direction = ParameterDirection.Input;
            HOURS_UPDATE_Param.Value = (from p in listOTWork select p.HOURS_UPDATE).ToArray();
            command.Parameters.Add(HOURS_UPDATE_Param);

            OracleParameter CREATE_TYPE_Param = new OracleParameter("CREATE_TYPE", OracleDbType.Varchar2);
            CREATE_TYPE_Param.Direction = ParameterDirection.Input;
            CREATE_TYPE_Param.Value = (from p in listOTWork select p.CREATE_TYPE).ToArray();
            command.Parameters.Add(CREATE_TYPE_Param);

            OracleParameter REF_WORK_ID_Param = new OracleParameter("REF_WORK_ID", OracleDbType.Varchar2);
            REF_WORK_ID_Param.Direction = ParameterDirection.Input;
            REF_WORK_ID_Param.Value = (from p in listOTWork select p.REF_WORK_ID).ToArray();
            command.Parameters.Add(REF_WORK_ID_Param);

            OracleParameter DEL_STATUS_Param = new OracleParameter("DEL_STATUS", OracleDbType.Varchar2);
            DEL_STATUS_Param.Direction = ParameterDirection.Input;
            DEL_STATUS_Param.Value = (from p in listOTWork select p.DEL_STATUS).ToArray();
            command.Parameters.Add(DEL_STATUS_Param);

            OracleParameter ACHIVE_DATE_Param = new OracleParameter("ACHIVE_DATE", OracleDbType.Date);
            ACHIVE_DATE_Param.Direction = ParameterDirection.Input;
            ACHIVE_DATE_Param.Value = (from p in listOTWork select p.ACHIVE_DATE).ToArray();
            command.Parameters.Add(ACHIVE_DATE_Param);



            command.ExecuteNonQuery();
            conn.Close();
            Comm.Logger.Info("All OT save.");

        }

        public void SaveApp1()
        {
            if (listApp1.Count() == 0) return;
            OracleConnection conn = new OracleConnection(context.Database.Connection.ConnectionString);
            OracleCommand command = new OracleCommand();
            command.Connection = conn;
            command.ArrayBindCount = listApp1.Count();
            command.CommandText = @"insert into OT_APP1 values(:APP1_ID,              
              :APP1_NO,              
              :CYCLE_ID,              
              :ORG_ID,              
              :STATUS,              
              :STATUS_FROM,              
              :APP_TYPE,              
              :CREATE_BY,              
              :CREATE_DATE,              
              :UPDATE_BY,
              :UPDATE_DATE,
              :VERSION_NUM,
              :WORKFLOW_ID)";
            conn.Open();
            OracleParameter APP1_ID_Param = new OracleParameter("APP1_ID", OracleDbType.Varchar2);
            APP1_ID_Param.Direction = ParameterDirection.Input;
            APP1_ID_Param.Value = (from p in listApp1 select p.APP1_ID).ToArray();
            command.Parameters.Add(APP1_ID_Param);

            OracleParameter APP1_NO_Param = new OracleParameter("APP1_NO", OracleDbType.Varchar2);
            APP1_NO_Param.Direction = ParameterDirection.Input;
            APP1_NO_Param.Value = (from p in listApp1 select p.APP1_NO).ToArray();
            command.Parameters.Add(APP1_NO_Param);

            OracleParameter CYCLE_ID_Param = new OracleParameter("CYCLE_ID", OracleDbType.Varchar2);
            CYCLE_ID_Param.Direction = ParameterDirection.Input;
            CYCLE_ID_Param.Value = (from p in listApp1 select p.CYCLE_ID).ToArray();
            command.Parameters.Add(CYCLE_ID_Param);

            OracleParameter ORG_ID_Param = new OracleParameter("ORG_ID", OracleDbType.Varchar2);
            ORG_ID_Param.Direction = ParameterDirection.Input;
            ORG_ID_Param.Value = (from p in listApp1 select p.ORG_ID).ToArray();
            command.Parameters.Add(ORG_ID_Param);

            OracleParameter STATUS_Param = new OracleParameter("STATUS", OracleDbType.Varchar2);
            STATUS_Param.Direction = ParameterDirection.Input;
            STATUS_Param.Value = (from p in listApp1 select p.STATUS).ToArray();
            command.Parameters.Add(STATUS_Param);

            OracleParameter STATUS_FROM_Param = new OracleParameter("STATUS_FROM", OracleDbType.Varchar2);
            STATUS_FROM_Param.Direction = ParameterDirection.Input;
            STATUS_FROM_Param.Value = (from p in listApp1 select p.STATUS_FROM).ToArray();
            command.Parameters.Add(STATUS_FROM_Param);

            OracleParameter APP_TYPE_Param = new OracleParameter("APP_TYPE", OracleDbType.Varchar2);
            APP_TYPE_Param.Direction = ParameterDirection.Input;
            APP_TYPE_Param.Value = (from p in listApp1 select p.APP_TYPE).ToArray();
            command.Parameters.Add(APP_TYPE_Param);

            OracleParameter CREATE_BY_Param = new OracleParameter("CREATE_BY", OracleDbType.Varchar2);
            CREATE_BY_Param.Direction = ParameterDirection.Input;
            CREATE_BY_Param.Value = (from p in listApp1 select p.CREATE_BY).ToArray();
            command.Parameters.Add(CREATE_BY_Param);

            OracleParameter CREATE_DATE_Param = new OracleParameter("CREATE_DATE", OracleDbType.Date);
            CREATE_DATE_Param.Direction = ParameterDirection.Input;
            CREATE_DATE_Param.Value = (from p in listApp1 select p.CREATE_DATE).ToArray();
            command.Parameters.Add(CREATE_DATE_Param);

            OracleParameter UPDATE_BY_Param = new OracleParameter("UPDATE_BY", OracleDbType.Varchar2);
            UPDATE_BY_Param.Direction = ParameterDirection.Input;
            UPDATE_BY_Param.Value = (from p in listApp1 select p.UPDATE_BY).ToArray();
            command.Parameters.Add(UPDATE_BY_Param);

            OracleParameter UPDATE_DATE_Param = new OracleParameter("UPDATE_DATE", OracleDbType.Date);
            UPDATE_DATE_Param.Direction = ParameterDirection.Input;
            UPDATE_DATE_Param.Value = (from p in listApp1 select p.UPDATE_DATE).ToArray();
            command.Parameters.Add(UPDATE_DATE_Param);

            OracleParameter VERSION_NUM_Param = new OracleParameter("VERSION_NUM", OracleDbType.Int16);
            VERSION_NUM_Param.Direction = ParameterDirection.Input;
            VERSION_NUM_Param.Value = (from p in listApp1 select p.VERSION_NUM).ToArray();
            command.Parameters.Add(VERSION_NUM_Param);

            OracleParameter WORKFLOW_ID_Param = new OracleParameter("WORKFLOW_ID", OracleDbType.Varchar2);
            WORKFLOW_ID_Param.Direction = ParameterDirection.Input;
            WORKFLOW_ID_Param.Value = (from p in listApp1 select p.WORKFLOW_ID).ToArray();
            command.Parameters.Add(WORKFLOW_ID_Param);


            command.ExecuteNonQuery();
            conn.Close();
            Comm.Logger.Info("All APP1 saved.");

        }

        public void SaveApp2()
        {
            if (listApp2.Count() == 0) return;
            OracleConnection conn = new OracleConnection(context.Database.Connection.ConnectionString);
            OracleCommand command = new OracleCommand();
            command.Connection = conn;
            command.ArrayBindCount = listApp2.Count();
            command.CommandText = @"insert into OT_APP2 values(:APP2_ID,              
              :APP2_NO,              
              :CYCLE_ID,              
              :ORG_ID,              
              :STATUS,              
              :STATUS_FROM,              
              :APP_TYPE,              
              :CREATE_BY,              
              :CREATE_DATE,              
              :UPDATE_BY,
              :UPDATE_DATE,
              :VERSION_NUM,
              :WORKFLOW_ID)";
            conn.Open();
            OracleParameter APP2_ID_Param = new OracleParameter("APP2_ID", OracleDbType.Varchar2);
            APP2_ID_Param.Direction = ParameterDirection.Input;
            APP2_ID_Param.Value = (from p in listApp2 select p.APP2_ID).ToArray();
            command.Parameters.Add(APP2_ID_Param);

            OracleParameter APP2_NO_Param = new OracleParameter("APP2_NO", OracleDbType.Varchar2);
            APP2_NO_Param.Direction = ParameterDirection.Input;
            APP2_NO_Param.Value = (from p in listApp2 select p.APP2_NO).ToArray();
            command.Parameters.Add(APP2_NO_Param);

            OracleParameter CYCLE_ID_Param = new OracleParameter("CYCLE_ID", OracleDbType.Varchar2);
            CYCLE_ID_Param.Direction = ParameterDirection.Input;
            CYCLE_ID_Param.Value = (from p in listApp2 select p.CYCLE_ID).ToArray();
            command.Parameters.Add(CYCLE_ID_Param);

            OracleParameter ORG_ID_Param = new OracleParameter("ORG_ID", OracleDbType.Varchar2);
            ORG_ID_Param.Direction = ParameterDirection.Input;
            ORG_ID_Param.Value = (from p in listApp2 select p.ORG_ID).ToArray();
            command.Parameters.Add(ORG_ID_Param);

            OracleParameter STATUS_Param = new OracleParameter("STATUS", OracleDbType.Varchar2);
            STATUS_Param.Direction = ParameterDirection.Input;
            STATUS_Param.Value = (from p in listApp2 select p.STATUS).ToArray();
            command.Parameters.Add(STATUS_Param);

            OracleParameter STATUS_FROM_Param = new OracleParameter("STATUS_FROM", OracleDbType.Varchar2);
            STATUS_FROM_Param.Direction = ParameterDirection.Input;
            STATUS_FROM_Param.Value = (from p in listApp2 select p.STATUS_FROM).ToArray();
            command.Parameters.Add(STATUS_FROM_Param);

            OracleParameter APP_TYPE_Param = new OracleParameter("APP_TYPE", OracleDbType.Varchar2);
            APP_TYPE_Param.Direction = ParameterDirection.Input;
            APP_TYPE_Param.Value = (from p in listApp2 select p.APP_TYPE).ToArray();
            command.Parameters.Add(APP_TYPE_Param);

            OracleParameter CREATE_BY_Param = new OracleParameter("CREATE_BY", OracleDbType.Varchar2);
            CREATE_BY_Param.Direction = ParameterDirection.Input;
            CREATE_BY_Param.Value = (from p in listApp2 select p.CREATE_BY).ToArray();
            command.Parameters.Add(CREATE_BY_Param);

            OracleParameter CREATE_DATE_Param = new OracleParameter("CREATE_DATE", OracleDbType.Date);
            CREATE_DATE_Param.Direction = ParameterDirection.Input;
            CREATE_DATE_Param.Value = (from p in listApp2 select p.CREATE_DATE).ToArray();
            command.Parameters.Add(CREATE_DATE_Param);

            OracleParameter UPDATE_BY_Param = new OracleParameter("UPDATE_BY", OracleDbType.Varchar2);
            UPDATE_BY_Param.Direction = ParameterDirection.Input;
            UPDATE_BY_Param.Value = (from p in listApp2 select p.UPDATE_BY).ToArray();
            command.Parameters.Add(UPDATE_BY_Param);

            OracleParameter UPDATE_DATE_Param = new OracleParameter("UPDATE_DATE", OracleDbType.Date);
            UPDATE_DATE_Param.Direction = ParameterDirection.Input;
            UPDATE_DATE_Param.Value = (from p in listApp2 select p.UPDATE_DATE).ToArray();
            command.Parameters.Add(UPDATE_DATE_Param);

            OracleParameter VERSION_NUM_Param = new OracleParameter("VERSION_NUM", OracleDbType.Int16);
            VERSION_NUM_Param.Direction = ParameterDirection.Input;
            VERSION_NUM_Param.Value = (from p in listApp2 select p.VERSION_NUM).ToArray();
            command.Parameters.Add(VERSION_NUM_Param);

            OracleParameter WORKFLOW_ID_Param = new OracleParameter("WORKFLOW_ID", OracleDbType.Varchar2);
            WORKFLOW_ID_Param.Direction = ParameterDirection.Input;
            WORKFLOW_ID_Param.Value = (from p in listApp2 select p.WORKFLOW_ID).ToArray();
            command.Parameters.Add(WORKFLOW_ID_Param);

            command.ExecuteNonQuery();
            conn.Close();
            Comm.Logger.Info("All APP2 Save!");

        }

        public void  SaveShift()
        {
            if (listShift.Count() == 0) return;
            OracleConnection conn = new OracleConnection(context.Database.Connection.ConnectionString);
            OracleCommand command = new OracleCommand();
            command.Connection = conn;
            command.ArrayBindCount = listShift.Count();
            command.CommandText = @"insert into OT_SHIFT values(
              :SHIFT_ID,              
              :SEQ_ID,              
              :SHIFT_NAME,              
              :STATUS,              
              :OT_TIME,              
              :CREATE_BY,              
              :CREATE_DATE,              
              :VERSION_NUM,              
              :OT_STYLE,              
              :OT_FIX)";
            conn.Open();
            OracleParameter SHIFT_IDParam = new OracleParameter("SHIFT_ID", OracleDbType.Varchar2);
            SHIFT_IDParam.Direction = ParameterDirection.Input;
            SHIFT_IDParam.Value = (from p in listShift select p.SHIFT_ID).ToArray();
            command.Parameters.Add(SHIFT_IDParam);

            OracleParameter SEQ_IDParam = new OracleParameter("SEQ_ID", OracleDbType.Decimal);
            SEQ_IDParam.Direction = ParameterDirection.Input;
            SEQ_IDParam.Value = (from p in listShift select p.SEQ_ID).ToArray();
            command.Parameters.Add(SEQ_IDParam);

            OracleParameter SHIFT_NAMEParam = new OracleParameter("SHIFT_NAME", OracleDbType.Varchar2);
            SHIFT_NAMEParam.Direction = ParameterDirection.Input;
            SHIFT_NAMEParam.Value = (from p in listShift select p.SHIFT_NAME).ToArray();
            command.Parameters.Add(SHIFT_NAMEParam);

            
            OracleParameter STATUSParam = new OracleParameter("STATUS", OracleDbType.Varchar2);
            STATUSParam.Direction = ParameterDirection.Input;
            STATUSParam.Value = (from p in listShift select p.STATUS).ToArray();
            command.Parameters.Add(STATUSParam);

            OracleParameter OT_TIMEParam = new OracleParameter("OT_TIME", OracleDbType.Decimal);
            OT_TIMEParam.Direction = ParameterDirection.Input;
            OT_TIMEParam.Value = (from p in listShift select p.OT_TIME).ToArray();
            command.Parameters.Add(OT_TIMEParam);

            OracleParameter CREATE_BYParam = new OracleParameter("CREATE_BY", OracleDbType.Varchar2);
            CREATE_BYParam.Direction = ParameterDirection.Input;
            CREATE_BYParam.Value = (from p in listShift select p.CREATE_BY).ToArray();
            command.Parameters.Add(CREATE_BYParam);

            OracleParameter CREATE_DATEParam = new OracleParameter("CREATE_DATE", OracleDbType.Date);
            CREATE_DATEParam.Direction = ParameterDirection.Input;
            CREATE_DATEParam.Value = (from p in listShift select p.CREATE_DATE).ToArray();
            command.Parameters.Add(CREATE_DATEParam);

            OracleParameter VERSION_NUMParam = new OracleParameter("VERSION_NUM", OracleDbType.Int16);
            VERSION_NUMParam.Direction = ParameterDirection.Input;
            VERSION_NUMParam.Value = (from p in listShift select p.VERSION_NUM).ToArray();
            command.Parameters.Add(VERSION_NUMParam);

            OracleParameter OT_STYLEParam = new OracleParameter("OT_STYLE", OracleDbType.Varchar2);
            OT_STYLEParam.Direction = ParameterDirection.Input;
            OT_STYLEParam.Value = (from p in listShift select p.OT_STYLE).ToArray();
            command.Parameters.Add(OT_STYLEParam);

            OracleParameter OT_FIXParam = new OracleParameter("OT_FIX", OracleDbType.Varchar2);
            OT_FIXParam.Direction = ParameterDirection.Input;
            OT_FIXParam.Value = (from p in listShift select p.OT_FIX).ToArray();
            command.Parameters.Add(OT_FIXParam);

            command.ExecuteNonQuery();
            conn.Close();
            Comm.Logger.Info("All shift saved.");

        }

        OT_APP3 CreateAPP3(OffsetItem item)
        {
            OT_APP3 app3 = new OT_APP3();
            app3.APP3_ID = Guid.NewGuid().ToString();
            app3.APP3_NO = SYS.getSeqNumByCode(ISysSeqService.SEQ_CODE_APP3);
            app3.STATUS = "0";
            if (item.Status.Equals("草稿"))
                app3.STATUS = "0";
            if (item.Status.Equals("待审"))
                app3.STATUS = "0";
            if (item.Status.Equals("已获批准"))
                app3.STATUS = "3";
            if (item.RemoveOffset_Hours != "")
                app3.OFFSET_TIME = decimal.Parse(item.RemoveOffset_Hours);
            else
                app3.OFFSET_TIME = 0;
            app3.ORG_ID = this.GroupId;
            app3.CREATE_BY = "System";
            app3.CREATE_DATE = DateTime.Now;
            app3.STATUS_FROM = "0";
            app3.APP_TYPE = "0";
            app3.CANCEL ="0";
            context.OT_APP3.Add(app3);
            context.SaveChanges();
            return app3;
          
        }

        bool CreateAudit(OTItem item,string app_id)
        {
            var q = from p in context.OT_EMP where p.EMP_NUMBER == item.App1 select p;
            if(q.FirstOrDefault()==null)
            {
                Comm.Logger.Warn(String.Format("Couldn't found the app1 work_number {0}, No Auditor created.",item.App1));
                return false;
            }
            OT_AUDIT audit1 = new OT_AUDIT();
            audit1.APP_ID = app_id;
            audit1.SEQ_ID = 1;
            audit1.AUDITER = q.FirstOrDefault().EMP_ID;
            audit1.CREATE_DATE = DateTime.Parse(item.Create_Ed);

            audit1.REMARKS = "从旧系统导入";
            audit1.VERSION_NUM = 0;
            audit1.CREATE_BY = "System";
            audit1.APP_TYPE = "0";//流程类别：0，加班主管及经理审批；1，加班部门经理及ht审核；2，补休申请；3，撤销补休申请
            audit1.STATUS = "1"; //?

            q = from p in context.OT_EMP where p.EMP_NUMBER == item.App2 select p;
            if (q.FirstOrDefault() == null)
            {
                Comm.Logger.Error(String.Format("Error--Couldn't found the app2 work_number  {0}", item.App2));
                return false;
            }
           

            OT_AUDIT audit2 = new OT_AUDIT();
            audit2.APP_ID = app_id;
            audit2.SEQ_ID = 2;
            audit2.AUDITER = q.FirstOrDefault().EMP_ID;
            audit2.CREATE_DATE = DateTime.Parse(item.Create_Ed);

            audit2.REMARKS = "从旧系统导入";
            audit2.VERSION_NUM = 0;
            audit2.CREATE_BY = "System";
            audit2.APP_TYPE = "1";//流程类别：0，加班主管及经理审批；1，加班部门经理及ht审核；2，补休申请；3，撤销补休申请
            audit2.STATUS = "1"; //?
            context.OT_AUDIT.Add(audit1);
            context.OT_AUDIT.Add(audit2);
            return true;

        }

        string NewOrSelectCycle(OTItem item)
        {
            DateTime started = DateTime.ParseExact(item.Cycle_StartEd, ONLYDATEFORMAT, ZH);
            DateTime ended = DateTime.ParseExact(item.Cycle_EndEd, ONLYDATEFORMAT, ZH);

            var q = context.OT_CYCLE
                .Where(x => DateTime.Compare(x.START_DATE.Value,started) == 0)
                .ToList()
                .Where(x => DateTime.Compare(x.END_DATE.Value,ended) == 0)
                .ToList();
            OT_CYCLE cyc = q.FirstOrDefault();
            if(cyc!=null)
            {
                return cyc.CYCLE_ID;
            }

            OT_CYCLE cycle = new OT_CYCLE();
            cycle.CYCLE_ID = Guid.NewGuid().ToString();
            cycle.START_DATE = DateTime.Parse(item.Cycle_StartEd);
            cycle.END_DATE = DateTime.Parse(item.Cycle_EndEd);
            cycle.CREATE_DATE = DateTime.Now;
            cycle.CREATE_BY = "System";
            context.OT_CYCLE.Add(cycle);
            context.SaveChanges();            
            return cycle.CYCLE_ID;
        }



        OT_APP2 CreateAPP2NoDB(OTItem item, string cycleId)
        {
            OT_APP2 app2 = new OT_APP2();
            app2.CYCLE_ID = cycleId;
            app2.ORG_ID = this.DeptId;
            app2.STATUS = "3";
            //审批状态：0，组别经理审批完成（新建）；1，部门经理审批；2，hr审核；3，审批完成
            //状态来源：0，下级上报（新建）；1，撤回；2，驳回；
            SysSeqServiceImpl sys = new SysSeqServiceImpl();

            app2.STATUS_FROM = "0";
            app2.CREATE_BY = "System";
            app2.CREATE_DATE = DateTime.Parse(item.Create_Ed);
            app2.APP2_NO = sys.getSeqNumByCode(SysSeqServiceImpl.SEQ_CODE_APP2);
            app2.APP2_ID = Guid.NewGuid().ToString();
            app2.APP_TYPE = "-1"; // 从旧系统导入
            app2.VERSION_NUM = 0;
            listApp2.Add(app2);
            return app2;
        }


        OT_APP2 CreateAPP2(OTItem item,string cycleId)
        {
            OT_APP2 app2 = new OT_APP2();
            app2.CYCLE_ID = cycleId;
            app2.ORG_ID = this.DeptId;
            app2.STATUS = "3";
            //审批状态：0，组别经理审批完成（新建）；1，部门经理审批；2，hr审核；3，审批完成
             //状态来源：0，下级上报（新建）；1，撤回；2，驳回；
            SysSeqServiceImpl sys=new SysSeqServiceImpl();

            app2.STATUS_FROM = "0";
            app2.CREATE_BY = "System";
            app2.CREATE_DATE = DateTime.Parse(item.Create_Ed);
            app2.APP2_NO = sys.getSeqNumByCode(SysSeqServiceImpl.SEQ_CODE_APP2);
            app2.APP2_ID = Guid.NewGuid().ToString();
            app2.APP_TYPE = "-1"; // 从旧系统导入
            app2.VERSION_NUM = 0;
            context.OT_APP2.Add(app2);
            return app2;
        }

        OT_APP1 CreateAPP1NoDB(OTItem item)
        {
            string cycleId = NewOrSelectCycle(item);
            OT_APP1 app1 = new OT_APP1();

            app1.CREATE_DATE = DateTime.Parse(item.Create_Ed);
            app1.CREATE_BY = "System";
            SysSeqServiceImpl sys = new SysSeqServiceImpl();

            //?? prefix "OB"
            app1.APP1_ID = Guid.NewGuid().ToString();
            app1.APP1_NO = sys.getSeqNumByCode(SysSeqServiceImpl.SEQ_CODE_APP1_PERSON);
            app1.APP_TYPE = "-1"; // 从旧系统导入
            app1.CYCLE_ID = cycleId;
            app1.ORG_ID = this.GroupId;

            app1.STATUS = "3"; //审批状态：0，草稿；1，一级审批；2，二级审批；3，审批完成	
            app1.STATUS_FROM = "0";//,状态来源：0，下级上报（新建）；1，撤回；2，驳回；	
            app1.VERSION_NUM = 0;//  

            listApp1.Add(app1);
            return app1;
        }


        OT_APP1 CreateAPP1(OTItem item)
        {
            //数据导入需要根据时间新建申报周期，导入时会根据时间创建或选择合适的申报周期
            string cycleId = NewOrSelectCycle(item);
            OT_APP1 app1 = new OT_APP1();

            app1.CREATE_DATE = DateTime.Parse(item.Create_Ed);
            app1.CREATE_BY = "System";
            SysSeqServiceImpl sys=new SysSeqServiceImpl();

            //?? prefix "OB"
            app1.APP1_ID = Guid.NewGuid().ToString();
            app1.APP1_NO = sys.getSeqNumByCode(SysSeqServiceImpl.SEQ_CODE_APP1_PERSON);
            app1.APP_TYPE = "-1"; // 从旧系统导入
            app1.CYCLE_ID = cycleId;
            app1.ORG_ID = this.GroupId;

            app1.STATUS = "3"; //审批状态：0，草稿；1，一级审批；2，二级审批；3，审批完成	
            app1.STATUS_FROM = "0";//,状态来源：0，下级上报（新建）；1，撤回；2，驳回；	
            app1.VERSION_NUM = 0;//


            context.OT_APP1.Add(app1);
            return app1;
        }

        /// <summary>
        /// process offset item
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>

        

        public bool ProcessOffset(OffsetItem item)
        {
            OT_OFFSET offset = new OT_OFFSET();
            offset.OFFSET_ID = Guid.NewGuid().ToString();
            CheckEmployee(item.Worker_CnName, item.Worker_Number, item.Worker_Group, item.Worker_Dept);
            offset.EMP_ID = this.EmpId;
            offset.ORG_ID = this.GroupId;
           
            offset.OFFSET_NO = item.Offset_Number; 
            if(offset.OFFSET_NO=="")offset.OFFSET_NO=SYS.getSeqNumByCode(SysSeqServiceImpl.SEQ_CODE_OFFSET);
            offset.OFFSET_REMARK = "从旧系统导入";
            if (item.Offset_Hours != "")
                offset.OFFSET_TIME = decimal.Parse(item.Offset_Hours);
            else
                offset.OFFSET_TIME = 0;
            if (item.Offset_StartEd != "")
            {
                if (item.Offset_EndTime != "")
                {
                    string strDate = item.Offset_StartEd.Substring(0, NOTIMEFORMAT.Length + 1).Trim() + " " + DateTime.FromOADate(double.Parse(item.Offset_StartTime)).ToString("h:mm tt");
                    offset.START_DATE = DateTime.ParseExact(strDate, DATEFORMAT, ZH);
                }
                else
                {
                    string strDate = item.Offset_StartEd.Substring(0, NOTIMEFORMAT.Length + 1).Trim();
                    offset.START_DATE = DateTime.ParseExact(strDate, NOTIMEFORMAT, ZH);

                }
                
            }
            if (item.Offset_EndEd != "")
            {
                if (item.Offset_EndTime != "")
                {
                    string strDate = item.Offset_EndEd.Substring(0, NOTIMEFORMAT.Length + 1).Trim() + " " + DateTime.FromOADate(double.Parse(item.Offset_EndTime)).ToString("h:mm tt");
                    offset.END_DATE = DateTime.ParseExact(strDate, DATEFORMAT, ZH);
                }
                else
                {
                    string strDate = item.Offset_EndEd.Substring(0, NOTIMEFORMAT.Length + 1).Trim();
                    offset.END_DATE = DateTime.ParseExact(strDate, NOTIMEFORMAT, ZH);
                }
            }
            offset.APP3_ID=CreateAPP3(item).APP3_ID;
            offset.CREATE_BY = "System";
            offset.CREATE_DATE = DateTime.ParseExact(item.OT_WorkEd.Substring(0,NOTIMEFORMAT.Length+1).Trim(), NOTIMEFORMAT, ZH);
            offset.EMP_ID = this.EmpId;
            offset.ORG_ID = this.GroupId;

            context.OT_OFFSET.Add(offset);
            context.SaveChanges();

           

             DateTime started = DateTime.ParseExact(item.OT_WorkEd, ONLYDATEFORMAT, ZH);

             var q = context.OT_WORK.Where(x => DateTime.Compare(x.START_DATE.Value, started) == 0).ToList()
                 .Where(x => x.EMP_ID == this.EmpId).ToList();

             OT_WORK ot_work=null;
            foreach(OT_WORK work_temp in q)
            {
                var q2=context.OT_WORK_OFFSET.Where(m=>m.WORK_ID==work_temp.WORK_ID).ToList();
                if (q2.FirstOrDefault() == null)
                    ot_work = work_temp;
            }
;
            if (ot_work != null)
            {
                OT_WORK_OFFSET existoffset = context.OT_WORK_OFFSET.Where(o => o.WORK_ID == ot_work.WORK_ID).FirstOrDefault();
                if(existoffset==null)
                { 
                    OT_WORK_OFFSET workoffset = new OT_WORK_OFFSET();
                    workoffset.WORK_OFFSET = Guid.NewGuid().ToString();
                    workoffset.OFFSET_ID = offset.OFFSET_ID;
                    workoffset.OFFSET_HOURS = offset.OFFSET_TIME;
                    workoffset.WORK_HOURS = ot_work.HOURS;
                    workoffset.WORK_ID = ot_work.WORK_ID;               
                    context.OT_WORK_OFFSET.Add(workoffset);
                    context.SaveChanges();
                }
                else
                {
                    existoffset.OFFSET_HOURS+=offset.OFFSET_TIME;
                    context.SaveChanges();
                    ExcelTool.WriteErrorOLEDB(item, "库存相同的记录,补休累加");
                   
                }
            }
            else
            {
                OTItem ot = new OTItem();
                ot.Create_Ed = item.OT_WorkEd;
                ot.OT_StartEd = item.OT_WorkEd;
                ot.OT_StartTime = item.OT_Work_StartTime;
                ot.OT_EndEd = item.OT_WorkEd;
                ot.OT_EndTime = item.OT_Work_EndTime;

                ot.OT_Hours = item.OT_Hours;
                ot.Reason = item.OT_Work_Comment;
                ot.Compensate_Rate = item.OT_CompensateRate;
                
                ot.Worker_CnName = item.Worker_CnName;
                ot.Worker_Dept = item.Worker_Dept;
                ot.Worker_Group = item.Worker_Group;
                ot.Worker_Number = item.Worker_Number;

                //Comm.Logger.Info("提取加班记录入库");
               
               // ProcessOTItem(ot);
                
                Comm.Logger.Error("没有在库中找到相关的加班记录");
                ExcelTool.WriteErrorOLEDB(item, "没有相关的加班记录");
                
               

            }
            Comm.Logger.Info("Save to DB ot_offset table success.");
            return true;
        }
    }
}

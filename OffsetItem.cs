using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Importer
{
    public class OffsetItem
    {

        public int OriginalRow;
        /// <summary>
        /// 申请单号-补休申请单号i.e.LB14A000001
        /// </summary>
        public string Apply_Number;

        /// <summary>
        /// 工号i.e.12124212
        /// </summary>
        public string Worker_Number;

        public string Worker_CnName;

        public string Worker_Dept;

        public string Worker_Group;

        public string CreateEd;

        /// <summary>
        /// 补休开始日期i.e.2014-10-15
        /// </summary>
        public string Offset_StartEd;

        /// <summary>
        /// 补休开始时间2014-10-15 06:00:00
        /// </summary>
        public string Offset_StartTime;


        public string Offset_EndEd;

        public string Offset_EndTime;

        /// <summary>
        /// 补休小时i.e.2
        /// </summary>
        public string Offset_Hours;

        /// <summary>
        /// 补休天数-为空，天数计算到小时数中.i.e.0
        /// </summary>
        public string Offset_Days;



        /// <summary>
        /// 一级审批人
        /// </summary>
        public string App1 { get; set; }

        /// <summary>
        /// 二级审批人-工号？
        /// </summary>
        public string App2 { get; set; }


        /// <summary>
        /// 扣减加班单单号（一条或多条）-唯一标识一次加班的号码的号码列表，
        /// 用英文斜杠“/”分隔，i.e.OB14A000001/OB14A000002
        /// </summary>
        public string OT_Numbers;

        /// <summary>
        /// 扣减补休小时（一条或多条）每条加班扣减单号对于加班单扣减小时数，
        /// 用英文斜杠“/”分隔顺序保持与扣减加班单单号一致.i.e.1.5/2.5
        /// </summary>
        public string OTOffset_Hours;

        /// <summary>
        /// 加班记录扣减补休小时
        /// </summary>
        public string RemoveOffset_Hours;




        /// <summary>
        /// 补休编号-如果没有，可以为空，系统导入时自动生成
        /// i.e.LD14A000001
        /// </summary>
        public string Offset_Number;


        /// <summary>
        /// 申报周期开始时间i.e.2014-10-01 00:00:00
        /// </summary>
        public string Cycle_StartEd;

        /// <summary>
        /// 申报周期结束时间i.e.2014-11-01 23:59:59
        /// </summary>
        public String Cycle_EndEd;

        /// <summary>
        /// 审批状态-只能是草稿、驳回、审批完成
        /// </summary>
        public String Status;


        /// <summary>
        /// 正常班次
        /// </summary>
        public string Shift_Id;

        //for Octorber version 
        /// <summary>
        /// 加班日期
        /// </summary>
        public string OT_WorkEd; 

        /// <summary>
        /// 加班开始时间.i.e.2015/10/4 
        /// </summary>
        public string OT_Work_StartTime;
        /// <summary>
        /// 加班结束时间i.e.9:00:00
        /// </summary>
        public string OT_Work_EndTime;
        /// <summary>
        /// 加班说明
        /// </summary>
        public string OT_Work_Comment;

        /// <summary>
        /// 加班倍数
        /// </summary>
        public string OT_CompensateRate;

        /// <summary>
        /// 加班小时
        /// </summary>
        public string OT_Hours;

    }
}

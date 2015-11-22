using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Importer
{
    public class OTItem
    {

        public int OriginalRow;
        /// <summary>
        /// 申报单号-如果一个申报单中有多条加班记录，则每条加班记录显示一行，申报单号相同，
        /// 如果旧系统没有此单号则新系统导入时根据申请的批次或申报周期自动生成i.e.OB14A000001
        /// </summary>
        public string OT_ApplyNumber { get; set; }

        /// <summary>
        /// 工号i.e.12124212
        /// </summary>
        public string Worker_Number { get; set; }

        /// <summary>
        /// 姓名
        /// </summary>
        public string Worker_CnName { get; set; }

        /// <summary>
        /// 部门i.e.ISD
        /// </summary>
        public string Worker_Dept { get; set; }

        /// <summary>
        /// 组别i.e.企业方案组
        /// </summary>
        public string Worker_Group { get; set; }

        /// <summary>
        /// 申报周期开始时间-由于新系统中的申报周期需要维护，所以就数据导入需要根据时间新建申报周期，
        /// 导入时会根据时间创建或选择合适的申报周期
        /// </summary>
        public string Cycle_StartEd { get; set; }


        /// <summary>
        /// 申报周期结束时间
        /// </summary>
        public string Cycle_EndEd { get; set; }

        /// <summary>
        /// 创建时间（申报时间）
        /// </summary>
        public string Create_Ed { get; set; }

        /// <summary>
        /// 加班开始日期
        /// </summary>
        public string OT_StartEd { get; set; }


        /// <summary>
        /// 加班开始时间
        /// </summary>
        public string OT_StartTime { get; set; }


        /// <summary>
        /// 加班结束日期i.e.2014-10-15

        /// </summary>
        public string OT_EndEd { get; set; }

        /// <summary>
        /// 加班结束时间i.e.2014-10-15 06:00:00 
        /// </summary>
        public string OT_EndTime { get; set; }


        /// <summary>
        /// 统计日期-加班日期
        /// </summary>
        public string Statistic_Date { get; set; }

        /// <summary>
        /// 加班小时
        /// </summary>
        public string OT_Hours { get; set; }

        /// <summary>
        /// 工资小时
        /// </summary>
        public string Pay_Hours { get; set; }

        /// <summary>
        /// 补休小时
        /// </summary>
        public string Offset_Hours { get; set;}


        /// <summary>
        /// 剩余补休小时i.e.2
        /// </summary>
        public string LeftOffset_Hour { get; set; }

        /// <summary>
        /// 改单小时
        /// </summary>
        public string LeftChange_Hour { get; set; }

        /// <summary>
        /// 加班种类-需要为导入数据确定一个加班种类
        /// </summary>
        public string OT_WorkType { get; set; }

       

        /// <summary>
        /// 加班原因-i.e.系统上线
        /// </summary>
        public string Reason { get; set; }

        /// <summary>
        /// 备注
        /// </summary>
        public string Comment { get; set; }

        /// <summary>
        /// 一级审批人
        /// </summary>
        public string App1 { get; set; }

        /// <summary>
        /// 二级审批人-工号？
        /// </summary>
        public string App2 { get; set; }

        /// <summary>
        /// 审批状态-固定为“审批完成”旧数据请考虑都走完审批再导入到新系统.i.e.审批完成
        /// </summary>
        public string Status { get; set; }


        /// <summary>
        /// 申报类别--个人申报还是集中申报 i.e.个人申报 
        /// </summary>
        public string Apply_Type { get; set; }

        /// <summary>
        /// 加班号-唯一标识一次加班的号码 i.e.OB14A000001
        /// </summary>
        public string OT_Number { get; set; }


        /// <summary>
        /// 考勤开始时间 i.e.2014-10-15 06:00:00
        /// </summary>
        public string Attendance_StartEd { get; set; }

        /// <summary>
        /// 考勤结束时间 i.e.2014-10-15 08:00:00
        /// </summary>
        public string Attendance_EndEd { get; set; }


        /// <summary>
        /// 加班班次,i.e.行政班,ADM,
        /// </summary>

        public string Shift_Id { get; set; }


        /// <summary>
        /// 加班补偿倍数,etc2
        /// </summary>
        public string Compensate_Rate { get; set; }


    }
}

using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace Importer
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        //[STAThread]
        static void Main(string[] args)
        {
            //Application.EnableVisualStyles();
            //Application.SetCompatibleTextRenderingDefault(false);
            //Application.Run(new MainFrame());
            Console.WriteLine("Import Tool V1.0.20151121");            
            Comm.InitLogger();
            ExcelTool.LoadDeptMap();
            CheckDBConnection();
            StartWork();
            Comm.Logger.Info("Work done!");
            Console.WriteLine("Work done!Press any key to quit.");
            Console.ReadKey();
        }

        static void  CheckDBConnection()
        {
                OTEntities ctx = new OTEntities();
    
                var q = from p in ctx.OT_APP1 select p;
                if (q.Count() >-1)
                {
                    Console.Write("db connection is good");
                }
                ctx.Database.Connection.Close();
                ctx.Dispose();            
            
        }



       


        static List<OTItem> validOTS=new List<OTItem>();
        static List<OTItem> ots = new List<OTItem>();
        static List<OffsetItem> offsets = new List<OffsetItem>();
        static List<OffsetItem> validOffsets=new List<OffsetItem>();
        static OTControler ctc = new OTControler();




        static bool CheckEmployInDBOrNot()
        {            
            Comm.Logger.Info("开始检查所有记录员工在库是否存在...");
            bool result = true;
            int i = 1;
            foreach (OTItem item in ots)
            {
                Console.Write(string.Format("检查OT记录 {0}...", i));
                bool ret=ctc.CheckEmployeeExist(item);
                if (!ret)
                {
                    ExcelTool.WriteErrorOLEDB(item, "员工不存在");
                }
                else
                    validOTS.Add(item);
               
                result = result & ret;
                i++;
                if (i % 100 == 0)
                    Console.Clear();
            }
            i = 1;
            foreach (OffsetItem item in offsets)
            {
                Console.Write(string.Format("检查OFFSET记录 {0}...", i));
                bool ret = ctc.CheckEmployeeExist(item);
                if (ret)
                    validOffsets.Add(item);
                else
                    ExcelTool.WriteErrorOLEDB(item, "员工不存在");

                result = result & ret;
                i++;
                if (i % 100 == 0)
                    Console.Clear();
            }
            if (!result)
            {

                StringBuilder content=new StringBuilder();
                foreach (KeyValuePair<string, string> item in Comm.NoExistEmp)
                {
                    content.AppendLine(item.Value);
                 }
                Comm.Logger.Error(content);
                Comm.Logger.Error(string.Format("以上共{0}个员工在库中不存在.", Comm.NoExistEmp.Count()));
               
            }

            return result;

        }

        static void JustSaveOTs()
        {
             int i = 1;


            Comm.Logger.Info("Begin save data to temp table...");
            foreach (OTItem item in ots)
            {

                Comm.Logger.Info(string.Format("正在处理第{0}条数据", i));
                ctc.SaveExcelOTs(item);
                i++;

            }
            //ctc.SaveChanges();
        }

        static void HandleOTs()
        {
           
            int i = 1;


            Comm.Logger.Info("开始导入加班数据...");
            foreach (OTItem item in validOTS)
            {

                Comm.Logger.Info(string.Format("正在处理第{0}条数据", i));
                if (!ctc.CheckEmployee(item.Worker_CnName, item.Worker_Number, item.Worker_Group, item.Worker_Dept))
                {
                    i++;
                    ExcelTool.WriteErrorOLEDB(item, "职位所在的部门与映射文件不一");
                    continue;
                }
                try
                {
                    ctc.ProcessOTItemNoDB(item);
                }
                catch(Exception e)
                {                    
                    Comm.Logger.Error(e.Message);
                    ExcelTool.WriteErrorOLEDB(item, e.Message);
                }

                i++;
                if (i % 5000 == 0)
                {
                    SaveList();
                    Console.Clear();
                }
            }


            SaveList();
        }

       static void SaveList()
        {
            ctc.SaveShift();
            ctc.SaveApp1();
            ctc.SaveApp2();
            ctc.SaveOTWork();
            OTControler.listShift.Clear();
            OTControler.listApp1.Clear();
            OTControler.listApp2.Clear();
            OTControler.listOTWork.Clear();

        }


        static void HandleOffsets()
        {
            int i = 1;
            OTControler ctc = new OTControler();
            Comm.Logger.Info("开始导入补休数据...");
            foreach (OffsetItem item in validOffsets)
            {
                Comm.Logger.Info("Process record " + i.ToString() + "...");
                if(i==75)
                {
                    Console.Write("75");
                }
                ctc.ProcessOffset(item);
                i++;
            }
        }

        static void LoadRecordsFromExcel()
        {
            string otfile = ConfigurationManager.AppSettings["ot.workfile"].ToString();
            string offsetfile = ConfigurationManager.AppSettings["ot.offsetfile"].ToString();
            Comm.Logger.Info(String.Format("开始读加班Excel文件{0}", otfile));

            ots = ExcelTool.LoadOTOctober(otfile);
            Comm.Logger.Info((String.Format("加班记录数：{0}", ots.Count())));
            Comm.Logger.Info(String.Format("开始读补休Excel文件{0}", offsetfile));
            offsets = ExcelTool.LoadOffsetOctober(offsetfile);
            Comm.Logger.Info((String.Format("补休记录数:{0}", offsets.Count())));
        }
        static void StartWork()
        {
            ExcelTool.OpenExcelOLEDB();
           
            LoadRecordsFromExcel();
           
            bool haveAll = CheckEmployInDBOrNot();
           // bool haveAll = true;

            string isTest = ConfigurationManager.AppSettings["test"].ToString();
            if (!isTest.ToLower().Equals("yes")&& !haveAll)
            {
                Comm.Logger.Info("请先手工在数据库创建再运行本程序!");
                return;
            }
            HandleOTs();
            HandleOffsets();
            Comm.Logger.Info("导入结束，开始确认补休信息....!");
            ctc.CheckOffsetIntegrated();
            ExcelTool.CloseExcel();

        }



       

        static void CheckOTRecords()
        {
            foreach(OTItem item in validOTS)
            {
                var q = from p in validOTS where p.OT_StartEd == item.OT_StartEd && p.Worker_Number==item.Worker_Number select p;
                if(q.Count()>1)
                {
                    Comm.Logger.Info(string.Format("duplicate records {0}", q.Count()));
                }
            }
        }
    }
}

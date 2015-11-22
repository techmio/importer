using System;
using System.Collections.Generic;
using System.Collections;
using System.Data;
using System.Data.Odbc;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop;
using Microsoft.Office.Interop.Excel;

namespace Importer
{
   public static class ExcelTool
    {
          public static Microsoft.Office.Interop.Excel.Workbook MyBook = null;
          public static Microsoft.Office.Interop.Excel.Application MyApp = null;
         
          public static Microsoft.Office.Interop.Excel.Worksheet MySheet = null;

          public static Application ErrApp =null; 
          public static Microsoft.Office.Interop.Excel.Workbook ErrBook =null;
          public static Microsoft.Office.Interop.Excel.Worksheet ErrSheet = null;

          

          public static List<DeptMapping> DeptList = new List<DeptMapping>();
          public static List<GroupMapping> GroupList = new List<GroupMapping>();
          public static string ConnStr="";
          public static OleDbConnection ErrorExcelConnection;
          
         public static void OpenExcel()
         {
             ErrApp=new Microsoft.Office.Interop.Excel.Application();
             MyApp = new Microsoft.Office.Interop.Excel.Application();
             ErrBook = ErrApp.Workbooks.Open(AppDomain.CurrentDomain.BaseDirectory + "\\error.xlsx");
       

         }

         public static void OpenExcelOLEDB()
         {
             string  connStr = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "\\error.xlsx";
             string extenedProperties = ";Extended Properties=\"Excel 8.0;HDR=NO;IMEX=3;READONLY=FALSE\"";
             connStr = connStr + extenedProperties;
             ConnStr = connStr;

         }


          public static void LoadOToledb(string filepath)
          {
             
          }



          public static System.Data.DataTable LoadExcelUsingOLEDB(string filePath)
          {
              string fileType = System.IO.Path.GetExtension(filePath);
              bool hasTitle = true;
              using (DataSet ds = new DataSet())
              {
                  string strCon = string.Format("Provider=Microsoft.ACE.OLEDB.{0}.0;" +
                                  "Extended Properties=\"Excel {1}.0;HDR={2};IMEX=1;\";" +
                                  "data source={3};",
                                  (fileType == ".xls" ? 4 : 12), (fileType == ".xls" ? 8 : 12), (hasTitle ? "Yes" : "NO"), filePath);
                  string strCom = " SELECT * FROM [Sheet1$]";
                  using (OleDbConnection myConn = new OleDbConnection(strCon))
                  using (OleDbDataAdapter myCommand = new OleDbDataAdapter(strCom, myConn))
                  {
                      myConn.Open();
                      myCommand.Fill(ds);
                  }
                  if (ds == null || ds.Tables.Count <= 0) return null;
                  return ds.Tables[0];
              }

          }

       public static void LoadDeptMap()
       {
             Application deptApp = new Microsoft.Office.Interop.Excel.Application();
             deptApp.Visible = false;
             Workbook deptBook = deptApp.Workbooks.Open(AppDomain.CurrentDomain.BaseDirectory + "\\部门组别映射表.xlsx");
             Worksheet deptSheet = (Worksheet)deptBook.Sheets[1];
             int lastRow = deptSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Row;
            for (int index = 2; index <= lastRow; index++)
            {
                if (!string.IsNullOrEmpty(deptSheet.Cells[index, 1].Value))
                {
                    DeptMapping dm = new DeptMapping();
                    dm.OldName = (string)deptSheet.Cells[index, 1].Value;
                    dm.oldAbbre = (string)deptSheet.Cells[index, 2].Value;
                    dm.NewName = (string)deptSheet.Cells[index, 2].Value;
                    dm.NewDeptNumber = ((int)deptSheet.Cells[index, 4].Value).ToString();
                    DeptList.Add(dm);
                }

            }

            deptSheet = (Worksheet)deptBook.Sheets[2];
            lastRow = deptSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Row;
            for (int index = 2; index <= lastRow; index++)
            {
                if (!string.IsNullOrEmpty(deptSheet.Cells[index, 1].Value))
                {
                    GroupMapping gm = new GroupMapping();
                    gm.OldDeptName = (string)deptSheet.Cells[index, 1].Value;
                    gm.OldGroupName = (string)deptSheet.Cells[index, 2].Value;
                    gm.NewGroupName = deptSheet.Cells[index, 3].Value == null ? "" : (string)deptSheet.Cells[index,3].Value;
                    gm.NewGroupNumber = deptSheet.Cells[index, 4].Value == null ? "" : ((int)deptSheet.Cells[index, 4].Value).ToString();
                    gm.NewDeptName = (string)deptSheet.Cells[index, 5].Value;
                    gm.NewDeptNumber = ((int)deptSheet.Cells[index, 6].Value).ToString();
                    GroupList.Add(gm);


                }
            }
            deptBook.Close();
            deptApp.Quit();
       }

       public static void GetExcelDeptGroupID(string oldDeptName,string oldGroupName,ref int deptNumber,ref int groupNumber,ref string deptName,ref string groupName)
       {
           var q = from p in DeptList where p.OldName == oldDeptName || p.oldAbbre == oldDeptName select p;
           if(q.FirstOrDefault()!=null)
           {
               deptNumber = int.Parse(q.FirstOrDefault().NewDeptNumber);
               oldDeptName = q.FirstOrDefault().OldName;
               GroupMapping gm =GroupList.Where(o=>o.OldDeptName==oldDeptName).ToList()
                   .Where(o=>o.OldGroupName==oldGroupName).ToList().FirstOrDefault();
               if(gm!=null)
               {
                   deptName = gm.NewDeptName;
                   groupName = gm.NewGroupName;
                   if (!string.IsNullOrEmpty(gm.NewGroupNumber) && gm.NewGroupNumber != "")
                       groupNumber = int.Parse(gm.NewGroupNumber);
               }
           }
           
       }

      

       /// <summary>
       /// 根据10月数据
       /// </summary>
       /// <param name="filePath"></param>
       /// <returns></returns>
       public static List<OffsetItem> LoadOffsetOctober(string filePath)
       {
           List<OffsetItem> ots = new List<OffsetItem>();
           int sheetNo = 1;
           if (MyApp == null) MyApp = new Microsoft.Office.Interop.Excel.Application();
           ExcelTool.MyApp.Visible = false;

           ExcelTool.MyBook = MyApp.Workbooks.Open(filePath);
           ExcelTool.MySheet = (Worksheet)MyBook.Sheets[sheetNo]; // Explict cast is not required here
           int lastRow = MySheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Row;
           for (int index = 2; index <= lastRow; index++)
           {
               System.Array MyValues = (System.Array)MySheet.get_Range("A" + index.ToString(), "V" + index.ToString()).Cells.Value;
               if (MyValues.GetValue(1, 2) == null && MyValues.GetValue(1, 3) == null)
               {
                   Console.WriteLine("OffsetItem no value on row " + index.ToString());
                   WriteErrorOLEDB(MyValues, string.Format("原{0},数据不完整",index), false);
                   continue;
               }

               ots.Add(new OffsetItem
                {
                    //根据10月数据import
                    OriginalRow=index,
                    Apply_Number = "",
                    Worker_Number = GetValue(MyValues,2), //B
                    Worker_CnName = GetValue(MyValues,3), //C
                    Worker_Dept = GetValue(MyValues, 19), //S 
                    Worker_Group = GetValue(MyValues,1), //A
                    CreateEd = GetValue(MyValues,13), //M column
                    Offset_StartEd = GetValue(MyValues,13), //M
                    Offset_StartTime = GetValue(MyValues,14),//N
                    Offset_EndEd = GetValue(MyValues,15), //M
                    Offset_EndTime = GetValue(MyValues,16),//P
                    RemoveOffset_Hours = GetValue(MyValues, 12), //L
                    Offset_Hours = GetValue(MyValues,18), //R
                    Offset_Days = "0", 
                    App1 = "",
                    App2 = "",

                    OT_Numbers = "", //no value
                   
                   

                    //如果没有，可以为空，系统导入时自动生成
                    Offset_Number =GetValue(MyValues,11), //K column
                    
                    Shift_Id=GetValue(MyValues,4),//D column
                    OT_WorkEd=GetValue(MyValues,5),//E
                    OT_Work_StartTime = GetValue(MyValues, 6), //F
                    OT_Work_EndTime=GetValue(MyValues,7), //G
                    OT_Work_Comment=GetValue(MyValues,8), //H
                    OT_Hours = GetValue(MyValues, 9), //I
                    OT_CompensateRate = GetValue(MyValues, 10), //J
                    Status = GetValue(MyValues,17) //Q column,已获批准 待审 草稿
                });




           }

           return ots;

       }

       public static void WriteErrorOLEDB(ProblemEmployee emp)
       {
           using (OleDbConnection connection = new OleDbConnection(ConnStr))
           {
               connection.Open();
               string mysql = "INSERT INTO [Sheet3$] VALUES (";
               StringBuilder sb = new StringBuilder(mysql);
               SetValueOnError(emp.Name, sb);
               SetValueOnError(emp.Number, sb);
               SetValueOnError(emp.OldDeptName, sb);
               SetValueOnError(emp.OldDeptId, sb);
               SetValueOnError(emp.OldGroupName, sb);
               SetValueOnError(emp.OldGroupId, sb);
               SetValueOnError(emp.DeptName, sb);
               SetValueOnError(emp.DeptId, sb);
               SetValueOnError(emp.GroupName,sb);
               SetValueOnError(emp.GroupId,sb);
               sb.AppendFormat("'{0}')", "新旧部门组别不一");

               OleDbCommand commande = new OleDbCommand(sb.ToString(), connection);
               commande.ExecuteNonQuery();
               connection.Close();
               connection.Dispose();
           }

       }

       public static void WriteError(ProblemEmployee emp)
       {
           ExcelTool.ErrSheet = (Worksheet)ErrBook.Sheets[3];
           int lastRow = ErrSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Row;
           int index = lastRow + 1;
           ErrSheet.Cells[index, 1] = emp.Name;
           ErrSheet.Cells[index, 2] = emp.Number;
           ErrSheet.Cells[index, 3] = emp.OldDeptName;
           ErrSheet.Cells[index, 4] = emp.OldDeptId;
           ErrSheet.Cells[index, 5] = emp.OldGroupName;
           ErrSheet.Cells[index, 6] = emp.OldGroupId;
           ErrSheet.Cells[index, 7] = emp.DeptName;
           ErrSheet.Cells[index, 8] = emp.DeptId;
           ErrSheet.Cells[index, 9] = emp.GroupName;
           ErrSheet.Cells[index, 10] = emp.GroupId;
           ErrSheet.Cells[index, 11] = "新旧部门组别不一";
           ErrBook.Save();
       }


       public static void WriteErrorOLEDB(ProblemOT pot)
       {
           using (OleDbConnection connection = new OleDbConnection(ConnStr))
           {
               connection.Open();
               string mysql = "INSERT INTO [Sheet4$] VALUES (";
               StringBuilder sb = new StringBuilder(mysql);
               SetValueOnError(pot.EmpName,sb);
               SetValueOnError(pot.EmpNumber,sb);
               SetValueOnError(pot.OT_Date,sb);
               SetValueOnError(pot.Apply_Number,sb);
               SetValueOnError(pot.OT_Hours,sb);
               SetValueOnError(pot.Offset_hours,sb);
               sb.Append("'')");
               OleDbCommand commande = new OleDbCommand(sb.ToString(), connection);
               commande.ExecuteNonQuery();
               connection.Close();
               connection.Dispose();
           }

       }


       public static void WriteError(ProblemOT pot)
       {
           ExcelTool.ErrSheet = (Worksheet)ErrBook.Sheets[4];
           int lastRow = ErrSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Row;
           int index = lastRow + 1;
           ErrSheet.Cells[index, 1] = pot.EmpName;
           ErrSheet.Cells[index, 2] = pot.EmpNumber;
           ErrSheet.Cells[index, 3] = pot.OT_Date;
           ErrSheet.Cells[index, 4] = pot.Apply_Number;
           ErrSheet.Cells[index, 5] = pot.OT_Hours;
           ErrSheet.Cells[index, 6] = pot.Offset_hours;

           ErrBook.Save();
       }


       public static void WriteErrorOLEDB(System.Array myVaues,string reason,bool isOT)
       {
           ExcelTool.OpenExcelOLEDB();
           using (OleDbConnection connection = new OleDbConnection(ConnStr))
           {
               connection.Open();
               string mysql="";

               if (isOT)
               {
                   mysql = "INSERT INTO [Sheet1$] VALUES (";
               }
               else
               {
                   mysql = "INSERT INTO [Sheet2$] VALUES (";
               }
               StringBuilder sb = new StringBuilder(mysql);
               for(int i=0;i<myVaues.Length;i++)
               {
                   try
                   {
                        sb.AppendFormat("{0},", myVaues.GetValue(i));
                   }
                   catch(Exception err)
                   {
                       sb.Append("'',");
                   }
               }         
               
               sb.AppendFormat("'{0}')", reason);
               OleDbCommand commande = new OleDbCommand(sb.ToString(), connection);
               commande.ExecuteNonQuery();
               connection.Close();
               connection.Dispose();
           }

       }


       public static void WriteError(System.Array MyValues, string reason,bool isOT)
       {
          
          
           if (isOT)
           {
               ExcelTool.ErrSheet = (Worksheet)ErrBook.Sheets[1];
               ErrApp.Visible = false;
               int lastRow = ErrSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Row;
               int index = lastRow + 1;
               //System.Array MyValues = (System.Array)MySheet.get_Range("A" + index.ToString(), "P" + index.ToString()).Cells.Value;
               for (int i = 1; i < 17; i++)
                   ErrSheet.Cells[index, i].value = MyValues.GetValue(1, i);
               ErrSheet.Cells[index, 17] = reason;
           }
           else
           {
               ExcelTool.ErrSheet = (Worksheet)ErrBook.Sheets[2];
               ErrApp.Visible = false;
               int lastRow = ErrSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Row;
               int index = lastRow + 1;
               for (int i = 1; i < 23; i++)
                   ErrSheet.Cells[index, i].value = MyValues.GetValue(1, i);
               ErrSheet.Cells[index, 23] = reason;


           }
           ErrBook.Save();

       }

       public static void SetValueOnError(string item, StringBuilder sb)
       {
            try {
                sb.AppendFormat("'{0}',", item.Replace("'", "''"));
            }
            catch(Exception e)
            {
                sb.Append("'',");
            }
       }

       public static void WriteErrorOLEDB(OffsetItem item, string reason)
       {

           using (OleDbConnection connection = new OleDbConnection(ConnStr))
           {
               connection.Open();
               string mysql = "INSERT INTO [Sheet2$] VALUES (";
               
               StringBuilder sb = new StringBuilder(mysql);

               SetValueOnError(item.Worker_Group,sb);
               SetValueOnError(item.Worker_Number, sb);
               SetValueOnError(item.Worker_CnName, sb);
               SetValueOnError(item.Shift_Id, sb);

               SetValueOnError(item.OT_WorkEd, sb);
               SetValueOnError(item.OT_Work_StartTime, sb);
               SetValueOnError(item.OT_Work_EndTime, sb);


               SetValueOnError(item.OT_Work_Comment, sb);
               SetValueOnError(item.OT_Hours, sb);
               SetValueOnError(item.OT_CompensateRate, sb);


               SetValueOnError(item.Offset_Number, sb);
               SetValueOnError(item.Offset_Hours, sb);

               SetValueOnError(item.Offset_StartEd, sb);
               SetValueOnError(item.Offset_StartTime, sb);

               SetValueOnError(item.Offset_EndEd, sb);
               SetValueOnError(item.Offset_EndTime, sb);
               SetValueOnError(item.Status, sb);
               SetValueOnError(item.Offset_Hours, sb);
               SetValueOnError(item.Worker_Dept, sb);
               SetValueOnError(null, sb);
               SetValueOnError(null, sb);
               SetValueOnError(null, sb);
               sb.AppendFormat("'原{0},{1}')", item.OriginalRow,reason);

               OleDbCommand commande = new OleDbCommand(sb.ToString(), connection);

               commande.ExecuteNonQuery();

               connection.Close();
               connection.Dispose();
           }

       }


       public static void  WriteError(OffsetItem item,string reason)
       {
           ErrApp.Visible = false;
          
           ErrSheet = (Worksheet)ErrBook.Sheets[2];
           int lastRow = ErrSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Row;
           int index = lastRow + 1;

           RowIndex = index;
           ColIndex = 1;
           
           SetValueOnError(item.Worker_Group);
           SetValueOnError(item.Worker_Number);
           SetValueOnError(item.Worker_CnName);
           SetValueOnError(item.Shift_Id);

           SetValueOnError(item.OT_WorkEd);
           SetValueOnError(item.OT_Work_StartTime);
           SetValueOnError(item.OT_Work_EndTime);


           SetValueOnError(item.OT_Work_Comment);
           SetValueOnError(item.OT_Hours);
           SetValueOnError(item.OT_CompensateRate);


           SetValueOnError(item.Offset_Number);
           SetValueOnError(item.Offset_Hours);

           SetValueOnError(item.Offset_StartEd);
           SetValueOnError(item.Offset_StartTime);

           SetValueOnError(item.Offset_EndEd);
           SetValueOnError(item.Offset_EndTime);
           SetValueOnError(item.Status);
           SetValueOnError(item.Offset_Hours);
           SetValueOnError(item.Worker_Dept);
           SetValueOnError(null);
           SetValueOnError(null);
           SetValueOnError(null);
           SetValueOnError(string.Format("原{0},", item.OriginalRow)+reason);
           ErrBook.Save();

       }

       static int ColIndex=1;

       public static void SetValueOnError(object value)
       {
           ErrSheet.Cells[RowIndex, ColIndex].Value=value;
           ColIndex++;
        }


       public static void WriteErrorOLEDB(OTItem item, string reason)
       {
           using (OleDbConnection connection = new OleDbConnection(ConnStr))
           {
               connection.Open();
               string mysql = "INSERT INTO [Sheet1$] VALUES (";
               StringBuilder sb = new StringBuilder(mysql);



               SetValueOnError(item.Worker_Number,sb);
               SetValueOnError(item.Worker_CnName, sb);
               SetValueOnError(item.Worker_Dept, sb);
               SetValueOnError(item.Worker_Group, sb);

               SetValueOnError(item.Cycle_StartEd, sb);
               SetValueOnError(item.Cycle_EndEd, sb);

               SetValueOnError(item.Create_Ed, sb);
               SetValueOnError(item.OT_StartEd, sb);

               SetValueOnError(item.OT_StartTime, sb);
               SetValueOnError(item.OT_EndTime, sb);

               SetValueOnError(item.OT_Hours, sb);
               SetValueOnError(item.Pay_Hours, sb);
               SetValueOnError(item.Offset_Hours, sb);

               SetValueOnError(item.Reason, sb);
               SetValueOnError(item.Shift_Id, sb);
               SetValueOnError(item.Compensate_Rate, sb);
               sb.AppendFormat("'原{0},{1}')", item.OriginalRow, reason);

               OleDbCommand commande = new OleDbCommand(sb.ToString(), connection);
               commande.ExecuteNonQuery();
               connection.Close();
               connection.Dispose();
           }
       }
       public static void WriteError(OTItem item,string reason)
       {
          
           ErrApp.Visible = false;
           //ErrBook = ErrApp.Workbooks.Open(AppDomain.CurrentDomain.BaseDirectory+"\\Error.xlsx");
           ErrSheet = (Worksheet)ErrBook.Sheets[1];
           int lastRow = ErrSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Row;
           int index = lastRow + 1;

           RowIndex = index;
           ColIndex = 1;

           SetValueOnError(item.Worker_Number);
           SetValueOnError(item.Worker_CnName);
           SetValueOnError(item.Worker_Dept); 
           SetValueOnError(item.Worker_Group);;

           SetValueOnError(item.Cycle_StartEd);;
           SetValueOnError(item.Cycle_EndEd);;

           SetValueOnError(item.Create_Ed);;
           SetValueOnError(item.OT_StartEd);;

           SetValueOnError(item.OT_StartTime);;
           SetValueOnError(item.OT_EndTime);;

           SetValueOnError(item.OT_Hours);;
           SetValueOnError(item.Pay_Hours);;
           SetValueOnError(item.Offset_Hours);;

           SetValueOnError(item.Reason);
           SetValueOnError(item.Shift_Id);
           SetValueOnError(item.Compensate_Rate);
           SetValueOnError(string.Format("原{0}", item.OriginalRow) + reason);
           ErrBook.Save();
                 

       }

       public static void CloseExcel()
       {
           if(ErrBook!=null)
             ErrBook.Close();

           if (MyBook != null)
            MyBook.Close();
           if(ErrApp!=null)
           ErrApp.Quit();
           if(MyApp!=null)
           MyApp.Quit();
           
       }

       
      
       public static List<OffsetItem> LoadOffset(string filePath)
       {
           List<OffsetItem> ots = new List<OffsetItem>();
           int sheetNo = 4;

           MyApp.Visible = false;

           MyBook = MyApp.Workbooks.Open(filePath);
           MySheet = (Worksheet)MyBook.Sheets[sheetNo]; // Explict cast is not required here
           int lastRow = MySheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Row;
           for (int index = 1; index <= lastRow; index++)
           {
                System.Array MyValues = (System.Array)MySheet.get_Range("A" + index.ToString(), "T" + index.ToString()).Cells.Value;
                if (MyValues.GetValue(1, 2) == null && MyValues.GetValue(1, 3) == null)
                {
                    Console.WriteLine("OffsetItem no value on row " + index.ToString());
                    WriteError(MyValues, string.Format(""), false);
                    continue;
                }
                RowIndex = 1;
                ots.Add(new OffsetItem
                {
                    Apply_Number = GetValue(MyValues),
                    Worker_Number = GetValue(MyValues),
                    Worker_CnName = GetValue(MyValues),
                    Worker_Dept = GetValue(MyValues),
                    Worker_Group = GetValue(MyValues),
                    CreateEd = GetValue(MyValues),
                    Offset_StartEd = GetValue(MyValues),
                    Offset_StartTime = GetValue(MyValues),
                    Offset_EndEd = GetValue(MyValues),
                    Offset_EndTime = GetValue(MyValues),
                    Offset_Hours = GetValue(MyValues),
                    Offset_Days = GetValue(MyValues),
                    App1 = GetValue(MyValues),
                    App2 = GetValue(MyValues),

                    OT_Numbers = GetValue(MyValues),
                    OTOffset_Hours = GetValue(MyValues),
                   
                    //如果没有，可以为空，系统导入时自动生成
                    Offset_Number = GetValue(MyValues),

                    Cycle_StartEd = GetValue(MyValues),
                    Cycle_EndEd = GetValue(MyValues),
                    Status = GetValue(MyValues)

                });

           }          
           return ots;
       }

       public static int RowIndex=1;


       public static string GetValue(Array MyValues)
       {
           String ret = "";
           ret=MyValues.GetValue(1, RowIndex) == null?"":MyValues.GetValue(1, RowIndex).ToString();
           RowIndex++;
           return ret;
       }

       public static string GetValue(Array MyValues,int rowIndex)
       {
           String ret = "";
           ret = MyValues.GetValue(1, rowIndex) == null ? "" : MyValues.GetValue(1, rowIndex).ToString();
           return ret;
       }

       public static string GetValue(Array MyValues, int rowIndex,bool isTime)
       {
           String ret = "";
           ret = MyValues.GetValue(1, rowIndex) == null ? "" : DateTime.FromOADate((double)MyValues.GetValue(1, rowIndex)).ToString("h:mm tt");
           return ret;
       }



       public static DateTime ToDateTime(double value)
       {
           string[] parts = value.ToString().Split(new char[] { '.' });

           int hours = Convert.ToInt32(parts[0]);
           int minutes = Convert.ToInt32(parts[1].Substring(0,9));

           if ((hours > 23) || (hours < 0))
           {
               throw new ArgumentOutOfRangeException("value", "decimal value must be no greater than 23.59 and no less than 0");
           }
           if ((minutes > 59) || (minutes < 0))
           {
               throw new ArgumentOutOfRangeException("value", "decimal value must be no greater than 23.59 and no less than 0");
           }
           DateTime d = new DateTime(1, 1, 1, hours, minutes, 0);
           return d;
       }

       /// <summary>
       /// 根据10月数据import
       /// </summary>
       /// <param name="filePath"></param>
       /// <returns></returns>
       public static List<OTItem> LoadOTOctober(string filePath)
       {
           List<OTItem> ots = new List<OTItem>();
           int sheetNo =1;
           if (MyApp == null) MyApp=new Microsoft.Office.Interop.Excel.Application();
           MyApp.Visible = false;
           MyBook = MyApp.Workbooks.Open(filePath);
           MySheet = (Worksheet)MyBook.Sheets[sheetNo]; // Explict cast is not required here
           int lastRow = MySheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Row;
           for (int index = 4; index <= lastRow; index++)
           {
               System.Array MyValues = (System.Array)MySheet.get_Range("A" + index.ToString(), "P" + index.ToString()).Cells.Value;
               if ((MyValues.GetValue(1, 1) == null && MyValues.GetValue(1, 2) == null) || MyValues.GetValue(1, 1).ToString().Trim()=="")
               {
                   Console.WriteLine("OTItem no value on row " + index.ToString());
                   WriteErrorOLEDB(MyValues, string.Format("原{0},数据不完整",index),true);
                   continue;
               }
               RowIndex = 1;

               ots.Add(new OTItem
               {
                   OriginalRow=index,
                   OT_ApplyNumber = "",
                   Worker_Number = GetValue(MyValues,1), //A
                   Worker_CnName = GetValue(MyValues,2), //B
                   Worker_Dept = GetValue(MyValues,3), //C
                   Worker_Group = GetValue(MyValues,4), //D

                   Cycle_StartEd = GetValue(MyValues,5), //E
                   Cycle_EndEd = GetValue(MyValues,6), //F
                   Create_Ed = GetValue(MyValues,7), //G

                   OT_StartEd = GetValue(MyValues,8), //H
                   OT_StartTime = GetValue(MyValues,9,true), //I
                   OT_EndEd = GetValue(MyValues,8), //H 
                   OT_EndTime = GetValue(MyValues,10,true),//J

                   Statistic_Date = "",
                   OT_Hours = GetValue(MyValues,11),//K
                   Pay_Hours = GetValue(MyValues,12),//L
                   Offset_Hours = GetValue(MyValues,13), //M
                   LeftOffset_Hour = "",
                   LeftChange_Hour = "",

                   //需要为导入数据确定一个加班种类,the value is -1
                   OT_WorkType ="-1",
                   Reason = GetValue(MyValues,14), //N
                   Comment = "",

                   App1 = "",
                   App2 = "",

                   //审批状态 固定为“审批完成”旧数据请考虑都走完审批再导入到新系统
                   Status = "审批完成",
                   Apply_Type ="",
                   //加班号,唯一标识一次加班的号码 "OD"
                   OT_Number ="",
                   Attendance_StartEd = "",
                   Attendance_EndEd = "",

                   Shift_Id = GetValue(MyValues,15), //o
                   Compensate_Rate = GetValue(MyValues,16) //P 
               });

           }


           return ots;

       }

       public static List<OTItem> LoadOT(string filePath)
       {
           List<OTItem> ots = new List<OTItem>();
           int sheetNo = 2;

           MyApp.Visible = false;

           MyBook = MyApp.Workbooks.Open(filePath);
           MySheet = (Worksheet)MyBook.Sheets[sheetNo]; // Explict cast is not required here
           int lastRow = MySheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Row;
           for (int index = 1; index <= lastRow; index++)
           {
               System.Array MyValues = (System.Array)MySheet.get_Range("A" + index.ToString(), "AD" + index.ToString()).Cells.Value;
               if (MyValues.GetValue(1, 2) == null && MyValues.GetValue(1, 3)==null)
               {
                   Console.WriteLine("OTItem no value on row " + index.ToString());
                   continue;
               }
               RowIndex = 1;
           
               ots.Add(new OTItem
               {
                   OT_ApplyNumber = GetValue(MyValues),
                   Worker_Number = GetValue(MyValues),
                   Worker_CnName = GetValue(MyValues),
                   Worker_Dept = GetValue(MyValues),
                   Worker_Group = GetValue(MyValues),

                   Cycle_StartEd = GetValue(MyValues),
                   Cycle_EndEd = GetValue(MyValues),
                   Create_Ed = GetValue(MyValues),

                   OT_StartEd = GetValue(MyValues),
                   OT_StartTime = GetValue(MyValues),
                   OT_EndEd = GetValue(MyValues),
                   OT_EndTime = GetValue(MyValues),

                   Statistic_Date = GetValue(MyValues),
                   OT_Hours = GetValue(MyValues),
                   Pay_Hours = GetValue(MyValues),
                   Offset_Hours = GetValue(MyValues),
                   LeftOffset_Hour = GetValue(MyValues),
                   LeftChange_Hour = GetValue(MyValues),
                  
                   //需要为导入数据确定一个加班种类,the value is -1
                   OT_WorkType = GetValue(MyValues),
                   Reason = GetValue(MyValues),
                   Comment = GetValue(MyValues),

                   App1 = GetValue(MyValues),
                   App2 = GetValue(MyValues),

                   //审批状态 固定为“审批完成”旧数据请考虑都走完审批再导入到新系统
                   Status = GetValue(MyValues),
                   Apply_Type = GetValue(MyValues),
                   //加班号,唯一标识一次加班的号码
                   OT_Number = GetValue(MyValues),
                   Attendance_StartEd = GetValue(MyValues),
                   Attendance_EndEd = GetValue(MyValues),

                   Shift_Id = GetValue(MyValues),
                   Compensate_Rate = GetValue(MyValues)
               });
               
           }

           return ots;
       }









        //加载Excel 
        public static DataSet LoadOTWithOLE(string filePath)
        {

                string strConn;
                strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";Extended Properties='Excel 10.0 XML'";
                OleDbConnection OleConn = new OleDbConnection(strConn);
                OleConn.Open();
                String sql = "SELECT * FROM  [Sheet2$]";//可是更改Sheet名称，比如sheet2，等等 

                OleDbDataAdapter OleDaExcel = new OleDbDataAdapter(sql, OleConn);
                DataSet OleDsExcle = new DataSet();
                OleDaExcel.Fill(OleDsExcle, "Sheet2");
                OleConn.Close();
                return OleDsExcle;

        }
    }
}

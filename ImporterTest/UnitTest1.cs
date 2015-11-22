using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Importer;
using System.Linq;



namespace ImporterTest
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestDetp1()
        {
            ExcelTool.LoadDeptMap();
            int deptNumber=0;
            int groupNumber=0;
            string deptName="";
            string groupName="";
            ExcelTool.GetExcelDeptGroupID("ENG", "设施维修组", ref deptNumber, ref groupNumber, ref deptName, ref  groupName);
            Assert.IsTrue(groupNumber == 106005);
            


        }

        [TestMethod]

        public void TestEF()
        {
            OTControler ctc = new OTControler();
            var q = from p in ctc.context.OT_APP1 select p;
            if (q.Count() > 0)
                Assert.IsTrue(true);
            else
                Assert.IsTrue(false);
        }

        [TestMethod]
        public void TestWriteOLEDBExcel()
        {
            Array myArray = new int[16]{ 1, 2, 3, 4, 5,6,7,8,9,10,11,12,13,14,15,16};

            ExcelTool.WriteErrorOLEDB(myArray, "test", true);
            ExcelTool.WriteErrorOLEDB(new OTItem(),"OTItem");
            ExcelTool.WriteErrorOLEDB(new OffsetItem(), "OffsetItem");
            ExcelTool.WriteErrorOLEDB(new ProblemOT());
            ExcelTool.WriteErrorOLEDB(new ProblemEmployee());


        }
    }
}

using log4net;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Importer
{
    public static class Comm
    {

        public static Dictionary<string, String> NoExistEmp = new Dictionary<string, string>();
        public static ILog Logger=null;
        public static ILog InitLogger()
        {            
           
               if(Comm.Logger==null)
               {
                   Comm.Logger= LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);                
                   log4net.Config.XmlConfigurator.ConfigureAndWatch(new FileInfo("log4net.config"));
               }
             return Comm.Logger;
        }
    }
}

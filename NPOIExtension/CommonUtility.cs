using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace NPOIExtension
{
    public class CommonUtility
    {
        public static DateTime ConvertToDateTime(string datetime)
        {
            DateTime dt = DateTime.Now;
            Regex r;
            //201403201826530021
            //是否是20140223000629格式
            if (datetime.Length == 14)
            {
                //20140223000629 @"^\d{14}"
                r = new Regex(@"^\d{14}");
                if (r.IsMatch(datetime))
                {
                    return Convert.ToDateTime(FormateDate(datetime));
                }
                else
                {
                    return DateTime.Now;
                }
            }
            else
                if (datetime.Length > 0 && datetime.Length < 14)
            {
                return Convert.ToDateTime(GuessFormateDate(datetime));
            }
            else
            {
                //正则表达式匹配
                //20140223 07:50:59 @"^\d{8} [0-9]+:[0-9]+:[0-9]+"
                r = new Regex(@"^\d{8} [0-9]+:[0-9]+:[0-9]+");
                if (r.IsMatch(datetime))
                {
                    string returnstring = datetime.Replace(":", "").Replace(" ", "");
                    return Convert.ToDateTime(FormateDate(returnstring));
                }
                else
                {
                    string returnstring = datetime.Substring(0, 14);
                    return Convert.ToDateTime(FormateDate(returnstring));
                }
            }
        }

        private static string FormateDate(string s)
        {
            char[] c = s.ToCharArray();
            string strYear = c[0].ToString() + c[1].ToString() + c[2].ToString() + c[3].ToString() + "-";
            string strMonth = c[4].ToString() + c[5].ToString() + "-";
            string strDay = c[6].ToString() + c[7].ToString() + " ";
            string strHour = c[8].ToString() + c[9].ToString() + ":";
            string strMin = c[10].ToString() + c[11].ToString() + ":";
            string strSec = c[12].ToString() + c[13].ToString();

            string resoult = strYear + strMonth + strDay + strHour + strMin + strSec;

            return resoult;
        }
        private static string GuessFormateDate(string s)
        {
            int tempStep = 14 - s.Length;
            string t = "";
            for (int i = 0; i <= tempStep; i++)
            {
                t += "1";
            }
            string g = s + t;

            return FormateDate(g);
        }
    }
}

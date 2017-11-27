using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace Hs.Tools
{
    public static class MyGrab
    {
        public static string GetContent(string url)
        {
            //抓
            XElement xml = XElement.Load(url);
            //取
            string txt = "------------ 数学大冒险 -------------" + "\r\n";
            var list = xml.Element("channel").Elements("item")
                .Select((m, index1) => txt += index1.ToString() + ":" + m.Element("title").Value + "\r\n")
                .Where((n, index2) => index2 < 5)
                .ToList();

            return txt;
        }
    }
}

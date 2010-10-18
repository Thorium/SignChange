using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.IO;
using System.Timers;
using SignatureChange.Properties;

namespace SignatureChange
{
    public partial class SignatureAdd
    {
        public static Random rnd = new System.Random();
        public string MkSign(string s, string tag, string filetype)
        {
            int tagpos = s.IndexOf("...");
            if (tagpos == -1)
            {
                tagpos = s.IndexOf("…");
                if (tagpos == -1) return s;
            }
            string begin = s.Substring(0, tagpos);
            string end=string.Empty;
            if (filetype == ".txt"){ 
                end = s.Substring(s.IndexOf("\n", tagpos));
            }
            if (filetype == ".rtf"){
                end = s.Substring(s.IndexOf(@"\par", tagpos));
            }
            if (filetype == ".htm"){
                end = s.Substring(s.IndexOf("</", tagpos));
            }
            return begin+"..."+tag+end;
        }

        public IEnumerable<string> ReadResources()
        {
            int i = 0;
            var res = string.Empty;
            while (res != null)
            {
                res = Resources.ResourceManager.GetString(i++.ToString());
                if(res!=null)yield return res;
            }
        }

        public void ChSign(string path)
        {
            var filut = new DirectoryInfo(path).GetFiles();

            var v = ReadResources().ToList();
            var c = v.Count();

            var sign = v.Skip(rnd.Next(c)).First();
            foreach (var f in filut)
            {
                var o = File.ReadAllText(f.FullName, Encoding.Default);
                var u = MkSign(o, sign, f.Extension);
                File.WriteAllText(f.FullName, u, Encoding.Default);
            }
        }
        public void ChSign(object sender, System.EventArgs e)
        {
            ChSign();
        }

        public void ChSign()
        {
            var u = Environment.GetEnvironmentVariable("USERPROFILE");
            var p1 = u + @"\AppData\Roaming\Microsoft\Signatures";
            var p2 = u + @"\Application Data\Microsoft\Signatures";

            if (Directory.Exists(p1)) ChSign(p1);
            if (Directory.Exists(p2)) ChSign(p2);

        }
        //Vaihdetaan 5min valein...
        Timer t = new Timer() { Interval = 300000 };

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            ChSign();
            t.Elapsed += ChSign;
            t.Start();

        }


        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            t.Stop();
            t.Dispose();
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}

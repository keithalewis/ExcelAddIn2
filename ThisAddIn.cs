using System;
using System.Net;

namespace ExcelAddInRibbon
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            var tpl = @"C:\Users\kalx\AppData\Roaming\Microsoft\Templates\test.xll";
            Uri uri = new Uri(@"https://xlladdins.com/64/test.xll");
            ServicePointManager.Expect100Continue = true;
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            WebClient webClient = new WebClient();
            webClient.DownloadFile(uri, tpl);
            Globals.ThisAddIn.Application.RegisterXLL(tpl);
            //Globals.ThisAddIn.Application.ActiveWorkbook.FollowHyperlink(xll);
            //var n = new Microsoft.Office.Tools.Excel.Workbook;
            //FollowHyperlink(xll);
            //Globals.ThisAddIn.Application.RegisterXLL(@"C:\Users\kalx\source\repos\xlladdins\xll_math\x64\Debug\xll_math.xll");
            //var result = Globals.ThisAddIn.Application.RegisterXLL(@"https://xlladdins.com/addins/64/test.xll");
            //result = result;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject() => new RibbonController();
        
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

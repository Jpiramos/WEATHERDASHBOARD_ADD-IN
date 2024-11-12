﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Windows.Forms;

namespace WeatherDashboardAddIn
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Garantir que a planilha está criada
            if (Globals.ThisAddIn.Application.Workbooks.Count == 0)
            {
                Globals.ThisAddIn.Application.Workbooks.Add();
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }
    }
}

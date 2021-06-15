using System;
using System.Globalization;
using Microsoft.Office.Core;

namespace Sobeys.ExcelAddIn
{
    public partial class ThisAddIn
    {
        private Ribbon _ribbon;
        private Bootstrapper _bootstrapper;

        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            _ribbon = new Ribbon();
            return _ribbon;
        }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            SetupLanguage();
            _bootstrapper = new Bootstrapper(_ribbon);
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            _bootstrapper.Dispose();
        }

        private void SetupLanguage()
        {
            var lcid = Globals.ThisAddIn.Application.LanguageSettings.LanguageID[MsoAppLanguageID.msoLanguageIDUI];
            var culture = new CultureInfo(lcid);
            System.Threading.Thread.CurrentThread.CurrentUICulture = culture;
            System.Threading.Thread.CurrentThread.CurrentCulture = culture;
        }

        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }
    }
}

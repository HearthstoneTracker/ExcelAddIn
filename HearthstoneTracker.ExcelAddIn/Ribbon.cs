using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace HearthstoneTracker.ExcelAddIn
{
    using HearthstoneTracker.ExcelAddIn;

    public partial class Ribbon
    {
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void importGames_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ImportGames();

        }

        private void importArenas_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ImportArenas();

        }
    }
}

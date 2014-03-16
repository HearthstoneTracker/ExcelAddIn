namespace HearthstoneTracker.ExcelAddIn
{
    using System;
    using System.Collections.Generic;
    using System.Data.Entity;
    using System.Linq;

    using HearthstoneTracker.ExcelAddIn.Model;

    using Office = Microsoft.Office.Core;
    using System.Globalization;
    using System.Reflection;
    using System.Windows.Forms;

    using Microsoft.Office.Interop.Excel;

    public partial class ThisAddIn
    {
        private static readonly GameFieldList gameFields;

        private static readonly ArenaFieldList arenaFields;

        private CultureInfo ciUS = new CultureInfo("en-US");

        static ThisAddIn()
        {
            gameFields = new GameFieldList()
                            {
                                new GameField("Id", g => g.Id.ToString()),
                                new GameField("Started", g => g.Started, "yyyy-mm-dd hh:mm"),
                                new GameField("Stopped", g => g.Stopped, "yyyy-mm-dd hh:mm"),
                                new GameField("Hero", g => g.Hero != null ? g.Hero.ClassName : ""),
                                new GameField("Opponent", g => g.OpponentHero != null ? g.OpponentHero.ClassName : ""),
                                new GameField("GameMode", g => ((GameMode)g.GameMode).ToString()),
                                new GameField("GoFirst", g => g.GoFirst),
                                new GameField("Turns", g => g.Turns, "0"),
                                new GameField("Victory", g => g.Victory),
                                new GameField("Server", g => g.Server),
                                new GameField("Deck_Id", g => g.Deck != null ? g.Deck.Id.ToString() : ""),
                                new GameField("Deck_Name", g => g.Deck != null ? g.Deck.Name : ""),
                                new GameField("Arena_Id", g => g.ArenaSession != null ? g.ArenaSession.Id.ToString() : "")
                            };
            arenaFields = new ArenaFieldList()
                              {
                                  new ArenaField("Id", a => a.Id.ToString()),
                                  new ArenaField("Hero", a => a.Hero != null ? a.Hero.ClassName : ""),
                                  new ArenaField("StartDate", a=>a.StartDate, "yyyy-mm-dd hh:mm"),
                                  new ArenaField("EndDate", a=>a.EndDate, "yyyy-mm-dd hh:mm"),
                                  new ArenaField("Wins", a=>a.Wins, "0"),
                                  new ArenaField("Losses", a=>a.Losses, "0"),
                                  new ArenaField("RewardGold", a=>a.RewardGold, "0"),
                                  new ArenaField("RewardDust", a=>a.RewardDust, "0"),
                                  new ArenaField("RewardPacks", a=>a.RewardPacks, "0"),
                              };
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        public void ImportGames()
        {
            string filename;
            var result = this.ChooseDatabase(out filename);
            if (result != DialogResult.OK)
            {
                return;
            }

            using (var model = OpenDatabase(filename))
            {
                var data = model.GameResults
                    .Include(x => x.Hero)
                    .Include(x => x.OpponentHero)
                    .Include(x => x.Deck)
                    .Include(x => x.ArenaSession)
                    .ToList();

                Import(gameFields, data);
            }
        }

        public void ImportArenas()
        {
            string filename;
            var result = this.ChooseDatabase(out filename);
            if (result != DialogResult.OK)
            {
                return;
            }

            using (var model = OpenDatabase(filename))
            {
                var data = model.ArenaSessions
                    .Include(x => x.Hero)
                    .ToList();

                Import(arenaFields, data);
            }
        }
        public void Import<T>(FieldList<T> fieldList, IList<T> data)
        {
            Application.ScreenUpdating = false;
            try
            {
                var sheet = (Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
                var activeCell = Globals.ThisAddIn.Application.ActiveCell;
                var activeRow = activeCell.Row;
                var activeCol = activeCell.Column;

                sheet.Range[sheet.Cells[activeRow, activeCol], sheet.Cells[sheet.Rows.Count - activeRow, activeCol + fieldList.Count]].ClearContents();

                var lastCol = SetRowValues(
                    sheet,
                    activeRow,
                    activeCol,
                    fieldList.Select(x => (object)x.Header).ToArray());

                int row = activeRow + 1;
                foreach (var item in data)
                {
                    SetRowValuesSmart(sheet, row, activeCol, item, fieldList);
                    row++;
                }

                for (int i = 0; i < fieldList.Count; i++)
                {
                    var range = sheet.Range[sheet.Cells[activeRow, activeCol + i], sheet.Cells[activeRow, activeCol + i]].EntireColumn;
                    range.GetType().InvokeMember("NumberFormat", BindingFlags.SetProperty, null, range, new object[] { fieldList[i].NumberFormat }, ciUS);
                    range.AutoFit();
                }
                var tableRange =
                    sheet.Range[sheet.Cells[activeRow, activeCol], sheet.Cells[sheet.Rows.Count - activeRow, activeCol + fieldList.Count - 1]];
                sheet.ListObjects.Add(XlListObjectSourceType.xlSrcRange, tableRange, false, XlYesNoGuess.xlYes).Name = "Games";
            }
            finally
            {
                Application.ScreenUpdating = true;
            }
        }

        private DialogResult ChooseDatabase(out string filename)
        {
            var dlg = new OpenFileDialog()
            {
                Title = "Choose Hearthstone Tracker database",
                Filter = "Database |*.sdf",
                DefaultExt = ".sdf"
            };

            var result = dlg.ShowDialog();
            filename = dlg.FileName;
            return result;
        }

        public HSDbContext OpenDatabase(string filename)
        {
            // string connectionString = "Provider=Microsoft.SQLSERVER.CE.OLEDB.4.0; Data Source=" + dlg.FileName;
            string connectionString = "Data Source=" + filename;
            return new HSDbContext(connectionString);
        }

        private void SetRowValuesSmart<T>(Worksheet sheet, int row, int colOffset, T item, params FieldList<T>[] fieldLists)
        {
            int lastCol = colOffset;
            foreach (var fields in fieldLists)
            {
                var fieldValues = fields.Select(x => x.Expression(item)).ToArray();
                SetRowValues(sheet, row, lastCol, fieldValues);
                lastCol += fieldValues.Length;
            }
        }

        private int SetRowValues(Worksheet sheet, int row, params object[] values)
        {
            return SetRowValues(sheet, row, 1, values);
        }

        private int SetRowValues(Worksheet sheet, int row, int startcolumn, params object[] values)
        {
            var cell = startcolumn;
            foreach (var value in values)
            {
                ((Range)sheet.Cells.Item[row, cell]).Value2 = value;
                cell++;
            }
            return cell - 1;
        }

        private string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
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

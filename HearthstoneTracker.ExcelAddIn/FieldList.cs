namespace HearthstoneTracker.ExcelAddIn
{
    using System;
    using System.Collections.Generic;

    using HearthstoneTracker.ExcelAddIn.Model;

    public class GameFieldList : FieldList<GameResult>
    {
    }

    public class ArenaFieldList : FieldList<ArenaSession>
    {
    }

    public class GameField : Field<GameResult>
    {
        public GameField(string header, Func<GameResult, object> expression, string numberFormat = "@")
            : base(header, expression, numberFormat)
        {
        }
    }

    public class ArenaField : Field<ArenaSession>
    {
        public ArenaField(string header, Func<ArenaSession, object> expression, string numberFormat = "@")
            : base(header, expression, numberFormat)
        {
        }
    }

    public class FieldList<T> : List<Field<T>>
    {
    }

    public class Field<T>
    {
        public string Header { get; set; }

        public string NumberFormat { get; set; }

        public Func<T, object> Expression { get; set; }

        public Field(string header, Func<T, object> expression, string numberFormat = "@")
        {
            this.Header = header;
            this.Expression = expression;
            this.NumberFormat = numberFormat;
        }
    }
}
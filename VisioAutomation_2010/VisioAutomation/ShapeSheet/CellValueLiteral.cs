﻿namespace VisioAutomation.ShapeSheet
{
    /// <summary>
    /// CellValueLiteral is used in those cases where you want a caller to provide an int, double, bool, or string value to be used as a formula
    /// In the case of string inputs, no special escaping of strings is performed. The caller must do any escaping.
    /// This struct allows you to have one method that takes multiple types as a parameter (via implicit conversion) and is
    /// similar to the use of XName in System.Linq.Xml.
    /// 
    /// IMPORTANT: The value stored is always a string. Any input will be converted to a string.
    /// </summary>
    public struct CellValueLiteral
    {
        private readonly string _stringval;

        public CellValueLiteral(string value)
        {
            this._stringval = value;
        }

        public CellValueLiteral(int value)
        {
            var culture = System.Globalization.CultureInfo.InvariantCulture;
            this._stringval = value.ToString(culture);
        }

        public CellValueLiteral(double value)
        {
            var culture = System.Globalization.CultureInfo.InvariantCulture;
            this._stringval = value.ToString(culture);
        }

        public CellValueLiteral(bool value)
        {
            this._stringval = value ? "TRUE" : "FALSE";
        }

        public string Value => this._stringval;

        public bool HasValue => this._stringval != null;

        public override string ToString() => this.Value;

        public static implicit operator CellValueLiteral(string value)
        {
            return new CellValueLiteral(value);
        }

        public static implicit operator CellValueLiteral(int value)
        {
            return new CellValueLiteral(value);
        }

        public static implicit operator CellValueLiteral(double value)
        {
            return new CellValueLiteral(value);
        }

        public static implicit operator CellValueLiteral(bool value)
        {
            return new CellValueLiteral(value);
        }
    }
}
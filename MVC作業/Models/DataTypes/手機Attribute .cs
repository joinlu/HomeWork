using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;

namespace MVC作業.Models.DataTypes
{
    public class 手機Attribute : DataTypeAttribute
    {
        public const string Mobile = @"\d{4}-\d{6}";

        public 手機Attribute() : base(DataType.Text)
        {
        }

        public override bool IsValid(object value)
        {
            if (value == null)
            {
                return true;
            }

            string str = (string)value;

            return Regex.IsMatch(str, Mobile);
        }
    }
}
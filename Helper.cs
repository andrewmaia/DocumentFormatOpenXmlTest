using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ConsoleApp1
{
    public class Helper
    {
        public static EnumValue<CellValues> ResolverTipo(Type t)
        {
            switch (t.Name.ToLower())
            {
                case "decimal":
                case "double":
                case "int64":
                case "int32":
                case "int":
                case "byte":
                    return CellValues.Number;
                case "datetime":
                    return CellValues.Date;
                case "bool":
                    return CellValues.Boolean;
                default:
                    return CellValues.String;
            }
        }
        public static string ResolverValor(object valor)
        {
            
            if (valor is decimal v)
                return v.ToString(new CultureInfo ("en"));

            if (valor is double d)
                return d.ToString(new CultureInfo ("en"));

            if (valor is float f)
                return f.ToString(new CultureInfo ("en"));                

            if (valor is DateTime dt)
                return dt.ToString(new CultureInfo ("en"));                 

            return valor.ToString();
        }
    }
}
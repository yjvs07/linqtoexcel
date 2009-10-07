using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Linq.Expressions;
using System.Threading;

namespace LinqToExcel
{
    public class ExcelQueryFactory
    {
        public string FileName { get; set; }
        public string WorksheetName { get; set; }
        private readonly Dictionary<string, string> _mapping = new Dictionary<string, string>();

        public ExcelQueryFactory()
        {
            WorksheetName = "Sheet1";
        }

        public ExcelQueryable<TSheetData> Worksheet<TSheetData>()
        {
            return Worksheet<TSheetData>(FileName, _mapping, WorksheetName);
        }

        public ExcelQueryable<Row> Worksheet()
        {
            return Worksheet<Row>(FileName, _mapping, WorksheetName);
        }

        public static ExcelQueryable<Row> Worksheet(string fileName)
        {
            return Worksheet<Row>(fileName, new Dictionary<string, string>(), "Sheet1");
        }

        public static ExcelQueryable<TSheetData> Worksheet<TSheetData>(string fileName)
        {
            return Worksheet<TSheetData>(fileName, new Dictionary<string, string>(), "Sheet1");
        }

        public static ExcelQueryable<TSheetData> Worksheet<TSheetData>(string fileName, string worksheetName)
        {
            return Worksheet<TSheetData>(fileName, new Dictionary<string, string>(), worksheetName);
        }

        public static ExcelQueryable<TSheetData> Worksheet<TSheetData>(string fileName,  Dictionary<string, string> mapping, string worksheetName)
        {
            if (fileName == null)
                throw new ArgumentNullException(fileName);


            return new ExcelQueryable<TSheetData>(fileName, mapping, worksheetName);
        }

        public void AddMapping<TSheetData>(Expression<Func<TSheetData, object>> property, string column)
        {
            //Get the property name
            var exp = (LambdaExpression)property;
            //exp.Body has 2 possible types
            //If the property type is native, then exp.Body == typeof(MemberExpression)
            //If the property type is not native, then exp.Body == typeof(UnaryExpression) in which 
            //case we can get the MemberExpression from its Operand property
            var mExp = (exp.Body.NodeType == ExpressionType.MemberAccess) ?
                (MemberExpression)exp.Body :
                (MemberExpression)((UnaryExpression)exp.Body).Operand;
            var propertyName = mExp.Member.Name;

            _mapping[propertyName] = column;
        }

        private string GetFileExtension(string fileName)
        {
            var afterLastPeriod = fileName.LastIndexOf(".") + 1;
            return fileName.Substring(afterLastPeriod, fileName.Length - afterLastPeriod);
        }
    }
}

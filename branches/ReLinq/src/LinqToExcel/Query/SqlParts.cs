using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LinqToExcel.Query
{
    public class SqlParts
    {
        public string Table { get; set;}
        private string _where;
        private string _orderBy;

        public override string ToString()
        {
            var sql = new StringBuilder();
            sql.AppendFormat("SELECT * FROM {0}", Table);
            if (!String.IsNullOrEmpty(_where))
                sql.AppendFormat("WHERE {0}", _where);
            if (!String.IsNullOrEmpty(_orderBy))
                sql.AppendFormat("ORDER BY {0}", _orderBy);
            return sql.ToString();
        }

    }
}

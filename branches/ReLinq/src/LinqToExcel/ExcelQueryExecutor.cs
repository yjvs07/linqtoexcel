using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using Remotion.Data.Linq;
using LinqToExcel.Query;
using System.IO;
using System.Data.OleDb;
using System.Data;
using Remotion.Logging;
using System.Reflection;
using Remotion.Data.Linq.Clauses.ResultOperators;
using Remotion.Data.Linq.Clauses;
using System.Collections;
using Remotion.Data.Linq.Clauses.StreamedData;

namespace LinqToExcel
{
    public class ExcelQueryExecutor : IQueryExecutor
    {
        private readonly ILog _log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);
        private readonly string _fileName;
        private readonly Dictionary<string, string> _columnMappings;
        private string _worksheetName;

        public ExcelQueryExecutor(string fileName,  Dictionary<string, string> columnMappings, string worksheetName)
        {
            _fileName = fileName;
            _columnMappings = columnMappings;
            _worksheetName = worksheetName;
        }

        /// <summary>
        /// Executes a query with a scalar result, i.e. a query that ends with a result operator such as Count, Sum, or Average.
        /// </summary>
        public T ExecuteScalar<T>(QueryModel queryModel)
        {
            return ExecuteSingle<T>(queryModel, false);
        }

        /// <summary>
        /// Executes a query with a single result object, i.e. a query that ends with a result operator such as First, Last, Single, Min, or Max.
        /// </summary>
        public T ExecuteSingle<T>(QueryModel queryModel, bool returnDefaultWhenEmpty)
        {
            return (returnDefaultWhenEmpty) ?
                ExecuteCollection<T>(queryModel).FirstOrDefault() :
                ExecuteCollection<T>(queryModel).First();
        }

        /// <summary>
        /// Executes a query with a collection result.
        /// </summary>
        public IEnumerable<T> ExecuteCollection<T>(QueryModel queryModel)
        {
            var connString = GetConnectionString();
            var sql = GetSqlStatement(queryModel);
            LogSqlStatement(connString, sql);

            CheckForNotSupportedResultOperators(queryModel.ResultOperators);

            var results = GetDataResults(connString, sql, queryModel);

            var projector = GetSelectProjector<T>(results.FirstOrDefault(), queryModel);

            //foreach (var resultOperator in queryModel.ResultOperators)
            //{
            //    var databaseResult = new StreamedSequence(results, new StreamedSequenceInfo(typeof(T), Expression.Constant(1)));
            //    var outputData = (StreamedSequence)resultOperator.ExecuteInMemory(databaseResult);
            //    results =  outputData.GetTypedSequence<T>();
            //}

            return results.Cast<T>(projector);
        }

        protected Func<object, T> GetSelectProjector<T>(object firstResult, QueryModel queryModel)
        {
            Func<object, T> projector = (result) => (T)Convert.ChangeType(result, typeof(T));
            if ((firstResult.GetType() != typeof(T)) &&
                (typeof(T) != typeof(int)) &&
                (typeof(T) != typeof(long)))
            {
                var proj = ProjectorBuildingExpressionTreeVisitor.BuildProjector<T>(queryModel.SelectClause.Selector);
                projector = (result) => proj(new ResultObjectMapping(queryModel.MainFromClause, result));
            }
            return projector;
        }

        protected SqlParts GetSqlStatement(QueryModel queryModel)
        {
            var sqlVisitor = new SqlGeneratorQueryModelVisitor(_worksheetName, _columnMappings);
            sqlVisitor.VisitQueryModel(queryModel);
            var sql = sqlVisitor.SqlStatement;

            var resultOperators = queryModel.ResultOperators;
            var sqlOperators = SqlResultOperators();
            foreach (var result in resultOperators)
                if (sqlOperators.ContainsKey(result.GetType()))
                    sqlOperators[result.GetType()](sql, queryModel);

            return sql;
        }

        private void Take(QueryModel query, SqlParts sql)
        {
            
        }

        private void UpdateAggregate(QueryModel queryModel, SqlParts sql, string aggregateName)
        {
            sql.Aggregate = string.Format("{0}({1})",
                aggregateName,
                GetResultColumnName(queryModel));
        }

        private string GetResultColumnName(QueryModel queryModel)
        {
            var mExp = queryModel.SelectClause.Selector as MemberExpression;
            return (_columnMappings.ContainsKey(mExp.Member.Name)) ?
                _columnMappings[mExp.Member.Name] :
                mExp.Member.Name;
        }

        protected Dictionary<Type, Action<SqlParts, QueryModel>> SqlResultOperators()
        {
            var dic = new Dictionary<Type, Action<SqlParts, QueryModel>>();
            dic[typeof(AverageResultOperator)] =
                (sql, query) => UpdateAggregate(query, sql, "AVG");
            dic[typeof(CountResultOperator)] =
                (sql, query) => sql.Aggregate = "COUNT(*)";
            dic[typeof(LongCountResultOperator)] =
                (sql, query) => sql.Aggregate = "COUNT(*)";
            dic[typeof(FirstResultOperator)] =
                (sql, query) => sql.Aggregate = "TOP 1 *";
            dic[typeof(MaxResultOperator)] =
                (sql, query) => UpdateAggregate(query, sql, "MAX");
            dic[typeof(MinResultOperator)] =
                (sql, query) => UpdateAggregate(query, sql, "MIN");
            dic[typeof(SumResultOperator)] =
                (sql, query) => UpdateAggregate(query, sql, "SUM");
            dic[typeof(TakeResultOperator)] =
                (sql, query) => Take(query, sql);
            return dic;
        }

        protected void CheckForNotSupportedResultOperators(IEnumerable<ResultOperatorBase> resultOperators)
        {
            var notSupportedList = new List<Type>
            {
                typeof(ContainsResultOperator),
            };

            var notSupported = (from x in resultOperators
                                where notSupportedList.Contains(x.GetType())
                                select x.GetType().Name.Replace("ResultOperator", ""))
                                .FirstOrDefault();

            if (notSupported != null)
                throw new NotSupportedException(
                    string.Format("LinqToExcel does not provide support for the {0}() method", notSupported));
        }

        /// <summary>
        /// Executes the sql query and returns the data results
        /// </summary>
        /// <typeparam name="T">Data type in the main from clause (queryModel.MainFromClause.ItemType)</typeparam>
        /// <param name="queryModel">Linq query model</param>
        protected IEnumerable<object> GetDataResults(string connectionString, SqlParts sql, QueryModel queryModel)
        {
            IEnumerable<object> results;
            using (var conn = new OleDbConnection(connectionString))
            using (var command = conn.CreateCommand())
            {
                conn.Open();
                command.CommandText = sql.ToString();
                command.Parameters.AddRange(sql.Parameters.ToArray());
                var data = command.ExecuteReader();

                var columns = GetColumnNames(data);
                if (columns.Count() == 1 && columns.First() == "Expr1000")
                    results = GetScalarResults(data);
                else if (queryModel.MainFromClause.ItemType == typeof(Row))
                    results = GetRowResults(data, columns);
                else
                    results = GetTypeResults(data, columns, queryModel);
            }
            return results;
        }

        private string GetConnectionString()
        {
            var connString = "";

            if (_fileName.ToLower().EndsWith("xlsx") ||
                _fileName.ToLower().EndsWith("xlsm"))
                connString = string.Format(
                    @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0 Xml;HDR=YES""",
                    _fileName);
            else if (_fileName.ToLower().EndsWith("xlsb"))
            {
                connString = string.Format(
                    @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0;HDR=YES""",
                    _fileName);
            }
            else if (_fileName.ToLower().EndsWith("csv"))
            {
                _worksheetName = Path.GetFileName(_fileName);
                connString = string.Format(
                        @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=""text;HDR=Yes;FMT=Delimited;""",
                        Path.GetDirectoryName(_fileName));
            }
            else
                connString = string.Format(
                    @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=""Excel 8.0;HDR=YES;""",
                    _fileName);

            return connString;
        }

        private IEnumerable<object> GetRowResults(IDataReader data, IEnumerable<string> columns)
        {
            var results = new List<object>();
            var columnIndexMapping = new Dictionary<string, int>();
            for (var i = 0; i < columns.Count(); i++)
                columnIndexMapping[columns.ElementAt(i)] = i;

            while (data.Read())
            {
                IList<Cell> cells = new List<Cell>();
                for (var i = 0; i < columns.Count(); i++)
                    cells.Add(new Cell(data[i]));
                results.CallMethod("Add", new Row(cells, columnIndexMapping));
            }
            return results.AsEnumerable();
        }

        private IEnumerable<object> GetTypeResults(IDataReader data, IEnumerable<string> columns, QueryModel queryModel)
        {
            var results = new List<object>();
            var fromType = queryModel.MainFromClause.ItemType;
            var props = fromType.GetProperties();
            while (data.Read())
            {
                var result = Activator.CreateInstance(fromType);
                foreach (var prop in props)
                {
                    var columnName = (_columnMappings.ContainsKey(prop.Name)) ?
                        _columnMappings[prop.Name] :
                        prop.Name;
                    if (columns.Contains(columnName))
                        result.SetProperty(prop.Name, Convert.ChangeType(data[columnName], prop.PropertyType));
                }
                results.Add(result);
            } 
            return results.AsEnumerable();
        }

        private IEnumerable<object> GetScalarResults(IDataReader data)
        {
            data.Read();
            return new List<object> { data[0] };
        }

        private void LogSqlStatement(string connectionString, SqlParts sqlParts)
        {
            if (_log.IsDebugEnabled)
            {
                _log.DebugFormat("Connection String: {0}", connectionString);
                _log.DebugFormat("SQL: {0}", sqlParts.ToString());
                for (var i = 0; i < sqlParts.Parameters.Count(); i++)
                    _log.DebugFormat("Param[{0}]: {1}", i, sqlParts.Parameters.ElementAt(i).Value);
            }
        }

        private IEnumerable<string> GetColumnNames(IDataReader data)
        {
            var columns = new List<string>();
            var sheetSchema = data.GetSchemaTable();
            foreach (DataRow row in sheetSchema.Rows)
                columns.Add(row["ColumnName"].ToString());

            //Log a warning for any property to column mappings that do not exist in the excel worksheet
            foreach (var kvp in _columnMappings)
            {
                if (!columns.Contains(kvp.Value))
                    _log.WarnFormat("'{0}' column that is mapped to the '{1}' property does not exist in the '{2}' worksheet",
                        kvp.Value, kvp.Key, _worksheetName);
            }

            return columns;
        }
    }
}

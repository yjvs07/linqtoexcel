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
            var postProcessing = new Dictionary<Type, Dictionary<bool, Func<ResultOperatorBase, IEnumerable<T>, T>>>();
            postProcessing[typeof(LastResultOperator)] = new Dictionary<bool, Func<ResultOperatorBase, IEnumerable<T>, T>>();
            postProcessing[typeof(LastResultOperator)][true] = (res, arg) => arg.LastOrDefault();
            postProcessing[typeof(LastResultOperator)][false] = (res, arg) => arg.Last();

            var preProcessing = new Dictionary<Type, Action<QueryModel, SqlParts>>();
            preProcessing[typeof (AverageResultOperator)] =
                (query, sql) => UpdateAggregate(query, sql, "AVG");
            preProcessing[typeof(CountResultOperator)] =
                (query, sql) => sql.Aggregate = "COUNT(*)";
            preProcessing[typeof(FirstResultOperator)] =
                (query, sql) => sql.Aggregate = "TOP 1 *";
            preProcessing[typeof(MaxResultOperator)] =
                (query, sql) => UpdateAggregate(query, sql, "MAX");
            preProcessing[typeof(MinResultOperator)] =
                (query, sql) => UpdateAggregate(query, sql, "MIN");
            preProcessing[typeof(SumResultOperator)] =
                (query, sql) => UpdateAggregate(query, sql, "SUM");
            preProcessing[typeof(TakeResultOperator)] = 
                Take;

            var connString = GetConnectionString();
            var sqlVisitor = new SqlGeneratorQueryModelVisitor(_worksheetName, _columnMappings);
            sqlVisitor.VisitQueryModel(queryModel);

            var resultOperator = queryModel.ResultOperators.FirstOrDefault();
            if (preProcessing.ContainsKey(resultOperator.GetType()))
                preProcessing[resultOperator.GetType()](queryModel, sqlVisitor.SqlStatement);

            LogSqlStatement(sqlVisitor.SqlStatement, sqlVisitor.SqlStatement.Parameters);

            IEnumerable<T> results;

            using (var conn = new OleDbConnection(connString))
            using (var command = conn.CreateCommand())
            {
                conn.Open();
                command.CommandText = sqlVisitor.SqlStatement;
                Console.WriteLine(command.CommandText);
                command.Parameters.AddRange(sqlVisitor.SqlStatement.Parameters.ToArray());
                var data = command.ExecuteReader();

                var columns = GetColumnNames(data);
                results = (queryModel.MainFromClause.ItemType == typeof(Row)) ?
                    GetRowResults<T>(data, columns) :
                    GetCustomTypeResults<T>(data, columns, queryModel);
            }

            
            if (postProcessing.ContainsKey(resultOperator.GetType()))
            {
                return postProcessing[resultOperator.GetType()][returnDefaultWhenEmpty](null, results);
            }
            else
            {
                return (returnDefaultWhenEmpty) ? 
                    results.FirstOrDefault() : 
                    results.First();
            }
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
            //if (queryModel.SelectClause.Selector.NodeType == ExpressionType.MemberAccess)
            //{
                var mExp = queryModel.SelectClause.Selector as MemberExpression;
                return (_columnMappings.ContainsKey(mExp.Member.Name)) ?
                    _columnMappings[mExp.Member.Name] :
                    mExp.Member.Name;
            //}
        }

        /// <summary>
        /// Executes a query with a collection result.
        /// </summary>
        public IEnumerable<T> ExecuteCollection<T>(QueryModel queryModel)
        {
            var resultOperator = queryModel.ResultOperators.FirstOrDefault();
            var postProcessing = new Dictionary<Type, Func<ResultOperatorBase, IEnumerable<T>, IEnumerable<T>>>();
            postProcessing[typeof (SkipResultOperator)] =
                (res, arg) => arg.Skip(res.Cast<SkipResultOperator>().GetConstantCount());

            var connString = GetConnectionString();

            var sql = new SqlGeneratorQueryModelVisitor(_worksheetName, _columnMappings);
            sql.VisitQueryModel(queryModel);
            LogSqlStatement(sql.SqlStatement, sql.SqlStatement.Parameters);

            IEnumerable<T> results;

            using (var conn = new OleDbConnection(connString))
            using (var command = conn.CreateCommand())
            {
                conn.Open();
                command.CommandText = sql.SqlStatement;
                Console.WriteLine(command.CommandText);
                command.Parameters.AddRange(sql.SqlStatement.Parameters.ToArray());
                var data = command.ExecuteReader();

                var columns = GetColumnNames(data);
                results = (queryModel.MainFromClause.ItemType == typeof(Row)) ? 
                    GetRowResults<T>(data, columns) : 
                    GetCustomTypeResults<T>(data, columns, queryModel);
            }

            if (resultOperator != null &&
                postProcessing.ContainsKey(resultOperator.GetType()))
                results = postProcessing[resultOperator.GetType()](resultOperator, results);

            return results;
        }

        private SqlParts GetSQLStatement(QueryModel queryModel)
        {
            var sql = new SqlGeneratorQueryModelVisitor(_worksheetName, _columnMappings);
            sql.VisitQueryModel(queryModel);
            LogSqlStatement(sql.SqlStatement, sql.SqlStatement.Parameters);
            return sql.SqlStatement;
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

            if (_log.IsDebugEnabled) _log.Debug("Connection String: " + connString);
            return connString;
        }

        private IEnumerable<T> GetRowResults<T>(IDataReader data, IEnumerable<string> columns)
        {
            var results = new List<T>();
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
            return results;
        }

        private IEnumerable<T> GetCustomTypeResults<T>(IDataReader data, IEnumerable<string> columns, QueryModel queryModel)
        {
            var results = new List<T>();
            var props = queryModel.MainFromClause.ItemType.GetProperties();
            while (data.Read())
            {
                var result = Activator.CreateInstance<T>();
                if (queryModel.ResultOperators.Count == 0 ||
                    queryModel.MainFromClause.ItemType == typeof(T))
                {
                    foreach (var prop in props)
                    {
                        var columnName = (_columnMappings.ContainsKey(prop.Name)) ?
                            _columnMappings[prop.Name] :
                            prop.Name;
                        if (columns.Contains(columnName))
                            result.SetProperty(prop.Name, Convert.ChangeType(data[columnName], prop.PropertyType));
                    }
                }
                else
                {
                    result = (T)Convert.ChangeType(data[0], typeof(T));
                }
                results.Add(result);
            }
            return results;
        }

        private void LogSqlStatement(string sqlString, IEnumerable<OleDbParameter> sqlParameters)
        {
            if (_log.IsDebugEnabled)
            {
                _log.Debug("SQL: " + sqlString);
                for (var i = 0; i < sqlParameters.Count(); i++)
                    _log.DebugFormat("Param[{0}]: {1}", i, sqlParameters.ElementAt(i).Value);
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

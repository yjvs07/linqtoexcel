using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using Remotion.Data.Linq;
using Remotion.Data.Linq.Clauses;
using System.Data.OleDb;
using Remotion.Data.Linq.Clauses.ResultOperators;
using System.Linq.Expressions;
using Remotion.Logging;
using System.Reflection;

namespace LinqToExcel.Query
{
    public class SqlGeneratorQueryModelVisitor : QueryModelVisitorBase
    {
        private readonly string _table;
        private string _aggregate;
        private string _where;
        private IEnumerable<OleDbParameter> _params = new List<OleDbParameter>();
        private string _orderBy;
        private Dictionary<string, string> _columnMappings;

        public SqlGeneratorQueryModelVisitor(string table, Dictionary<string, string> columnMappings)
        {
            _table = table;
            _columnMappings = columnMappings;
            _aggregate = "*";
        }

        public override void VisitGroupJoinClause(GroupJoinClause groupJoinClause, QueryModel queryModel, int index)
        {
            throw new NotSupportedException("Group join clause is not supported");
        }

        public override void VisitJoinClause(JoinClause joinClause, QueryModel queryModel, int index)
        {
            throw new NotSupportedException("Join clause is not supported");
        }

        public override void VisitQueryModel(QueryModel queryModel)
        {
            queryModel.SelectClause.Accept(this, queryModel);
            queryModel.MainFromClause.Accept(this, queryModel);
            VisitBodyClauses(queryModel.BodyClauses, queryModel);
            VisitResultOperators(queryModel.ResultOperators, queryModel);
        }

        public override void VisitWhereClause(WhereClause whereClause, QueryModel queryModel, int index)
        {
            var where = new WhereClauseExpressionTreeVisitor(queryModel.MainFromClause.ItemType, _columnMappings);
            where.Visit(whereClause.Predicate);
            _where = where.WhereClause;
            _params = where.Params;

            base.VisitWhereClause(whereClause, queryModel, index);
        }

        public override void VisitResultOperator(ResultOperatorBase resultOperator, QueryModel queryModel, int index)
        {
            if (resultOperator is CountResultOperator)
                _aggregate = "COUNT(*)";
            else if (resultOperator is FirstResultOperator)
                _aggregate = "TOP 1 *";
            else if (resultOperator is SumResultOperator)
            {
                if (queryModel.SelectClause.Selector.NodeType == ExpressionType.MemberAccess)
                {
                    var mExp = queryModel.SelectClause.Selector as MemberExpression;
                    var columnName = (_columnMappings.ContainsKey(mExp.Member.Name)) ?
                        _columnMappings[mExp.Member.Name] :
                        mExp.Member.Name;
                    _aggregate = string.Format("SUM({0})", columnName);
                }

            }
            base.VisitResultOperator(resultOperator, queryModel, index);
        }

        public string GetSqlString()
        {
            var sql = new StringBuilder();
            sql.AppendFormat("SELECT {0} FROM [{1}$]", _aggregate, _table);
            if (_table.EndsWith("csv"))
                sql.Replace("$", "");
            if (!String.IsNullOrEmpty(_where))
                sql.AppendFormat(" WHERE {0}", _where);
            if (!String.IsNullOrEmpty(_orderBy))
                sql.AppendFormat(" ORDER BY {0}", _orderBy);
            return sql.ToString();
        }

        public IEnumerable<OleDbParameter> SqlParams
        {
            get { return _params; }
        }
    }
}

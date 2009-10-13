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
using Remotion.Collections;

namespace LinqToExcel.Query
{
    public class SqlGeneratorQueryModelVisitor : QueryModelVisitorBase
    {
        public SqlParts SqlStatement { get; protected set; }
        private Dictionary<string, string> _columnMappings;

        public SqlGeneratorQueryModelVisitor(string table, Dictionary<string, string> columnMappings)
        {
            SqlStatement = new SqlParts();
            SqlStatement.Table = string.Format("[{0}$]", table);
            _columnMappings = columnMappings;
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
            SqlStatement.Where = where.WhereClause;
            SqlStatement.Parameters = where.Params;

            base.VisitWhereClause(whereClause, queryModel, index);
        }

        public override void VisitResultOperator(ResultOperatorBase resultOperator, QueryModel queryModel, int index)
        {
            if (resultOperator is AverageResultOperator)
            {
                if (queryModel.SelectClause.Selector.NodeType == ExpressionType.MemberAccess)
                {
                    var mExp = queryModel.SelectClause.Selector as MemberExpression;
                    var columnName = (_columnMappings.ContainsKey(mExp.Member.Name)) ?
                        _columnMappings[mExp.Member.Name] :
                        mExp.Member.Name;
                    SqlStatement.Aggregate = string.Format("AVG({0})", columnName);
                }
            }
            else if (resultOperator is CastResultOperator)
                throw new NotImplementedException();
            else if (resultOperator is ContainsResultOperator)
                throw new NotImplementedException();
            else if (resultOperator is CountResultOperator)
                SqlStatement.Aggregate = "COUNT(*)";
            else if (resultOperator is DefaultIfEmptyResultOperator)
                throw new NotImplementedException();
            else if (resultOperator is DistinctResultOperator)
                throw new NotImplementedException();
            else if (resultOperator is ExceptResultOperator)
                throw new NotImplementedException();
            else if (resultOperator is FirstResultOperator)
                SqlStatement.Aggregate = "TOP 1 *";
            else if (resultOperator is GroupResultOperator)
                throw new NotImplementedException();
            else if (resultOperator is IntersectResultOperator)
                throw new NotImplementedException();
            else if (resultOperator is LastResultOperator)
            {
                //do nothing now
            }
            else if (resultOperator is LongCountResultOperator)
                throw new NotImplementedException();
            else if (resultOperator is MaxResultOperator)
            {
                if (queryModel.SelectClause.Selector.NodeType == ExpressionType.MemberAccess)
                {
                    var mExp = queryModel.SelectClause.Selector as MemberExpression;
                    var columnName = (_columnMappings.ContainsKey(mExp.Member.Name)) ?
                        _columnMappings[mExp.Member.Name] :
                        mExp.Member.Name;
                    SqlStatement.Aggregate = string.Format("MAX({0})", columnName);
                }
            }
            else if (resultOperator is MinResultOperator)
            {
                if (queryModel.SelectClause.Selector.NodeType == ExpressionType.MemberAccess)
                {
                    var mExp = queryModel.SelectClause.Selector as MemberExpression;
                    var columnName = (_columnMappings.ContainsKey(mExp.Member.Name)) ?
                        _columnMappings[mExp.Member.Name] :
                        mExp.Member.Name;
                    SqlStatement.Aggregate = string.Format("MIN({0})", columnName);
                }
            }
            else if (resultOperator is OfTypeResultOperator)
                throw new NotImplementedException();
            else if (resultOperator is ReverseResultOperator)
                throw new NotImplementedException();
            else if (resultOperator is SingleResultOperator)
                throw new NotImplementedException();
            else if (resultOperator is SkipResultOperator)
                throw new NotImplementedException();
            else if (resultOperator is SumResultOperator)
            {
                if (queryModel.SelectClause.Selector.NodeType == ExpressionType.MemberAccess)
                {
                    var mExp = queryModel.SelectClause.Selector as MemberExpression;
                    var columnName = (_columnMappings.ContainsKey(mExp.Member.Name)) ?
                        _columnMappings[mExp.Member.Name] :
                        mExp.Member.Name;
                    SqlStatement.Aggregate = string.Format("SUM({0})", columnName);
                }
            }
            else if (resultOperator is TakeResultOperator)
                throw new NotImplementedException();
            else if (resultOperator is UnionResultOperator)
                throw new NotImplementedException();
            base.VisitResultOperator(resultOperator, queryModel, index);
        }

        protected override void VisitBodyClauses(ObservableCollection<IBodyClause> bodyClauses, QueryModel queryModel)
        {
            var orderClause = bodyClauses.FirstOrDefault() as OrderByClause;
            if (orderClause != null)
            {
                var mExp = orderClause.Orderings.First().Expression as MemberExpression;
                var columnName = (_columnMappings.ContainsKey(mExp.Member.Name)) ?
                    _columnMappings[mExp.Member.Name] :
                    mExp.Member.Name;
                SqlStatement.OrderBy = columnName;
                var orderDirection = orderClause.Orderings.First().OrderingDirection;
                SqlStatement.OrderByAsc = (orderDirection == OrderingDirection.Asc) ? true : false;
            }
            base.VisitBodyClauses(bodyClauses, queryModel);
        }
    }
}

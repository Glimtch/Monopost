using System;
using System.Collections.Generic;
using System.Text;

namespace SB.Monopost.Shared
{
    public class ApcCellsCalculator
    {
        static void RecalculateFreeCells(IOrganizationService service, Guid apcId)
        {
            var ordersQuery = new QueryExpression("sb_order")
            {
                ColumnSet = new ColumnSet("sb_orderid")
            };
            ordersQuery.Criteria.AddCondition("sb_shipto_apcid", ConditionOperator.Equal, apcId);
            ordersQuery.Criteria.AddCondition("sb_orderstatuscode", ConditionOperator.Equal, 110000001);
            var orders = service.RetrieveMultiple(ordersQuery).Entities;

            int deliveriesCount = 0;
            foreach (var order in orders)
            {
                var deliveryQuery = new QueryExpression("sb_delivery")
                {
                    ColumnSet = new ColumnSet(false)
                };

                FilterExpression filter = new FilterExpression(LogicalOperator.And);
                FilterExpression filter1 = new FilterExpression(LogicalOperator.And);
                filter1.Conditions.Add(new ConditionExpression("sb_orderid", ConditionOperator.Equal, order.Id));
                FilterExpression filter2 = new FilterExpression(LogicalOperator.Or);
                filter2.Conditions.Add(new ConditionExpression("sb_deliverystatuscode", ConditionOperator.Equal, 110000003));
                filter2.Conditions.Add(new ConditionExpression("sb_deliverystatuscode", ConditionOperator.Equal, 110000004));
                filter.AddFilter(filter1);
                filter.AddFilter(filter2);
                deliveryQuery.Criteria = filter;

                deliveriesCount += service.RetrieveMultiple(deliveryQuery).Entities.Count;
            }

            var apc = service.Retrieve("sb_automatedpostalcenter", apcId, new ColumnSet("sb_cellscount"));

            apc["sb_freecellscount"] = apc.GetAttributeValue<int>("sb_cellscount") - deliveriesCount;

            service.Update(apc);
        }
    }
}

using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Extensions;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SB.DriverOrder.Messages
{
    public class PostUpdate : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            var context = serviceProvider.Get<IExecutionContext>();
            var service = serviceProvider.GetOrganizationService(context.UserId);

            try
            {
                var target = (Entity)context.InputParameters["Target"];
                
                switch(target.GetAttributeValue<OptionSetValue>("sb_deliverystatuscode").Value)
                {
                    case 110000001:

                        var query = new QueryExpression("sb_delivery")
                        {
                            ColumnSet = new ColumnSet(false)
                        };
                        query.Criteria.AddCondition("sb_driverorderid", ConditionOperator.Equal, target.Id);
                        var deliveries = service.RetrieveMultiple(query).Entities;



                        foreach (var delivery in deliveries)
                        {
                            delivery["sb_deliverystatuscode"] = new OptionSetValue(110000003);
                            delivery["sb_dateshipped"] = DateTime.Now;
                            service.Update(delivery);
                        }

                        var driverOrder = service.Retrieve("sb_driverorder", target.Id, new ColumnSet(true));

                        var vehicleId = driverOrder.GetAttributeValue<EntityReference>("sb_deliveryvehicleid").Id;
                        var vehicle = service.Retrieve("sb_deliveryvehicle", vehicleId, new ColumnSet());
                        vehicle["sb_statuscode"] = new OptionSetValue(110000001);
                        service.Update(vehicle);

                        RecalculateFreeCells(service, driverOrder.GetAttributeValue<EntityReference>("sb_shipto_apcid").Id);

                        break;

                    case 110000002:

                        var deliveryQuery = new QueryExpression("sb_delivery")
                        {
                            ColumnSet = new ColumnSet(false)
                        };
                        deliveryQuery.Criteria.AddCondition("sb_driverorderid", ConditionOperator.Equal, target.Id);
                        var deliveryEntities = service.RetrieveMultiple(deliveryQuery).Entities;

                        var order = service.Retrieve("sb_driverorder", target.Id, new ColumnSet(true));
                        var freeCells = GetFreeCellNumbers(service, order.GetAttributeValue<EntityReference>("sb_shipto_apcid").Id);

                        foreach (var delivery in deliveryEntities)
                        {
                            delivery["sb_deliverystatuscode"] = new OptionSetValue(110000004);
                            delivery["sb_datedelivered"] = DateTime.Now;
                            delivery["sb_dateexpiring"] = DateTime.Now.AddDays(4);
                            delivery["sb_deliverto_apccell"] = freeCells.First();
                            freeCells.RemoveAt(0);
                            service.Update(delivery);
                        }

                        var deliveryVehicleId = order.GetAttributeValue<EntityReference>("sb_deliveryvehicleid").Id;
                        var deliveryVehicle = service.Retrieve("sb_deliveryvehicle", deliveryVehicleId, new ColumnSet());
                        deliveryVehicle["sb_statuscode"] = new OptionSetValue(110000000);
                        service.Update(deliveryVehicle);

                        break;
                }
            }
            catch (Exception e)
            {
                throw new InvalidPluginExecutionException(e.Message);
            }
        }

        static List<int> GetFreeCellNumbers(IOrganizationService service, Guid apcId)
        {
            var freeCells = new List<int>();

            var ordersQuery = new QueryExpression("sb_order")
            {
                ColumnSet = new ColumnSet("sb_orderid")
            };
            ordersQuery.Criteria.AddCondition("sb_shipto_apcid", ConditionOperator.Equal, apcId);
            ordersQuery.Criteria.AddCondition("sb_orderstatuscode", ConditionOperator.Equal, 110000001);
            var orders = service.RetrieveMultiple(ordersQuery).Entities;

            var deliveries = new List<Entity>();
            foreach (var order in orders)
            {
                var deliveryQuery = new QueryExpression("sb_delivery")
                {
                    ColumnSet = new ColumnSet("sb_deliverto_apccell")
                };

                FilterExpression filter = new FilterExpression(LogicalOperator.And);
                FilterExpression filter1 = new FilterExpression(LogicalOperator.And);
                filter1.Conditions.Add(new ConditionExpression("sb_orderid", ConditionOperator.Equal, order.Id));
                filter1.Conditions.Add(new ConditionExpression("sb_deliverto_apccell", ConditionOperator.NotNull));
                FilterExpression filter2 = new FilterExpression(LogicalOperator.Or);
                filter2.Conditions.Add(new ConditionExpression("sb_deliverystatuscode", ConditionOperator.Equal, 110000003));
                filter2.Conditions.Add(new ConditionExpression("sb_deliverystatuscode", ConditionOperator.Equal, 110000004));
                filter.AddFilter(filter1);
                filter.AddFilter(filter2);
                deliveryQuery.Criteria = filter;

                deliveries.AddRange(service.RetrieveMultiple(deliveryQuery).Entities);
            }

            var cellNums = new List<int>();

            foreach(var delivery in deliveries)
            {
                cellNums.Add(delivery.GetAttributeValue<int>("sb_deliverto_apccell"));
            }

            var apc = service.Retrieve("sb_automatedpostalcenter", apcId, new ColumnSet("sb_freecellscount"));

            for (int i = 1; i < apc.GetAttributeValue<int>("sb_freecellscount"); i++)
            {
                if (cellNums.Contains(i))
                    continue;
                freeCells.Add(i);
            }

            return freeCells;
        }

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

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
    public class PostCreate : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            var context = serviceProvider.Get<IExecutionContext>();
            var service = serviceProvider.GetOrganizationService(context.UserId);

            try
            {
                var target = (Entity)context.InputParameters["Target"];

                var deliveries = GetTopPriorDeliveries(service, target);
                if (deliveries.Count() == 0)
                    throw new InvalidPluginExecutionException("There are no single deliveries by this route");
                foreach (var delivery in deliveries)
                {
                    delivery["sb_driverorderid"] = new EntityReference("sb_driverorder", target.Id);

                    service.Update(delivery);
                }
                
                var postOfficeId = target.GetAttributeValue<EntityReference>("sb_shipfrom_postofficeid").Id;
                var apcId = target.GetAttributeValue<EntityReference>("sb_shipto_apcid").Id;

                var postOfficeName = service.Retrieve("sb_postoffice", postOfficeId, new ColumnSet("sb_name"))["sb_name"];
                var apcName = service.Retrieve("sb_automatedpostalcenter", apcId, new ColumnSet("sb_name"))["sb_name"];
                var date = DateTime.Now.ToString("dd/MM/yyyy");

                target["sb_name"] = $"{ postOfficeName } - { apcName } - {date}";

                service.Update(target);
                
            }
            catch (Exception e)
            {
                throw new InvalidPluginExecutionException(e.Message);
            }
        }

        static IEnumerable<Entity> GetTopPriorDeliveries(IOrganizationService service, Entity driverOrder)
        {
            var fromPostOfficeId = driverOrder.GetAttributeValue<EntityReference>("sb_shipfrom_postofficeid").Id;
            var toApcId = driverOrder.GetAttributeValue<EntityReference>("sb_shipto_apcid").Id;

            var ordersQuery = new QueryExpression("sb_order")
            {
                ColumnSet = new ColumnSet("sb_orderid", "sb_shipfrom_postofficeid", "sb_shipto_apcid", "sb_orderstatuscode")
            };
            ordersQuery.Criteria.AddCondition("sb_shipfrom_postofficeid", ConditionOperator.Equal, fromPostOfficeId);
            ordersQuery.Criteria.AddCondition("sb_shipto_apcid", ConditionOperator.Equal, toApcId);
            ordersQuery.Criteria.AddCondition("sb_orderstatuscode", ConditionOperator.Equal, 110000001);

            var orders = service.RetrieveMultiple(ordersQuery).Entities;
            var deliveries = new List<Entity>();
            foreach (var order in orders)
            {
                var deliveryQuery = new QueryExpression("sb_delivery")
                {
                    ColumnSet = new ColumnSet(false)
                };
                deliveryQuery.Criteria.AddCondition("sb_orderid", ConditionOperator.Equal, order.Id);
                deliveryQuery.Criteria.AddCondition("sb_deliverystatuscode", ConditionOperator.Equal, 110000001);

                deliveries.AddRange(service.RetrieveMultiple(deliveryQuery).Entities);
            }

            var vehicleId = driverOrder.GetAttributeValue<EntityReference>("sb_deliveryvehicleid").Id;
            var vehicle = service.Retrieve("sb_deliveryvehicle", vehicleId, new ColumnSet("sb_capacity"));
            var apc = service.Retrieve("sb_automatedpostalcenter", toApcId, new ColumnSet("sb_freecellscount"));
            var maxDeliveries = Math.Min(vehicle.GetAttributeValue<int>("sb_capacity"), apc.GetAttributeValue<int>("sb_freecellscount"));
            return deliveries.OrderBy(d => d.GetAttributeValue<DateTime>("sb_datecreated")).Take(maxDeliveries);
        }
    }
}

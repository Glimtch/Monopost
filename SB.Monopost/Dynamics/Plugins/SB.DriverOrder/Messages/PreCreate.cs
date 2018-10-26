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
    public class PreCreate : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            var context = serviceProvider.Get<IExecutionContext>();
            var service = serviceProvider.GetOrganizationService(context.UserId);

            var target = (Entity)context.InputParameters["Target"];

            var vehicleId = target.GetAttributeValue<EntityReference>("sb_deliveryvehicleid").Id;
            var vehicle = service.Retrieve("sb_deliveryvehicle", vehicleId, new ColumnSet("sb_statuscode"));
            if (vehicle.GetAttributeValue<OptionSetValue>("sb_statuscode").Value != 110000000)
                throw new InvalidPluginExecutionException("Vehicle chosen is not active /n");
        }
    }
}

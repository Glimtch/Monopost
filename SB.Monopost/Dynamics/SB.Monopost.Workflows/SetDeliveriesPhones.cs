using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SB.Monopost.Workflows
{
    public sealed partial class SetDeliveriesPhones : CodeActivity
    {
        [RequiredArgument]
        [Input("Order Recipient")]
        [ReferenceTarget("contact")]
        public InArgument<EntityReference> Contact { get; set; }

        [RequiredArgument]
        [Input("Order")]
        [ReferenceTarget("sb_order")]
        public InArgument<EntityReference> Order { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            IWorkflowContext workflowContext = context.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = context.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(workflowContext.InitiatingUserId);

            var query = new QueryExpression("sb_delivery")
            {
                ColumnSet = new ColumnSet(false)
            };
            query.Criteria.AddCondition("sb_orderid", ConditionOperator.Equal, Order.Get(context).Id);

            var deliveries = service.RetrieveMultiple(query).Entities;

            var contactId = Contact.Get(context).Id;
            var contact = service.Retrieve("contact", contactId, new ColumnSet("telephone1"));

            foreach (var delivery in deliveries)
            {
                delivery["sb_phonetonotificate"] = contact.GetAttributeValue<string>("telephone1");
                
                service.Update(delivery);
            }
        }
    }
}

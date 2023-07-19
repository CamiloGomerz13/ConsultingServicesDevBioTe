using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BT.PLG.OpportunityCloseAsWon
{
    public class ValidateTitleCredentialToClose : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            {
                Entity entity = (Entity)context.InputParameters["Target"];
                IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
                try
                {
                    tracingService.Trace("1");
                    Entity opportunityClose = service.Retrieve("opportunityclose", entity.Id, new ColumnSet("opportunityid"));
                    tracingService.Trace("2");

                    Entity opportunity = service.Retrieve("opportunity", opportunityClose.GetAttributeValue<EntityReference>("opportunityid").Id, new ColumnSet("name"));
                    tracingService.Trace("3");

                    bool hasTitleCredential = true;

                    QueryExpression query = new QueryExpression("bio_trainingattendee");
                    query.ColumnSet.AddColumns("bio_opportunity", "bio_contactid");
                    query.Criteria.AddCondition("bio_opportunity", ConditionOperator.Equal, opportunity.Id);
                    EntityCollection attendees = service.RetrieveMultiple(query);
                    tracingService.Trace("4");

                    foreach (var item in attendees.Entities)
                    {
                        tracingService.Trace("5");

                        Entity contact = service.Retrieve("contact", item.GetAttributeValue<EntityReference>("bio_contactid").Id, new ColumnSet("jobtitle"));
                        tracingService.Trace("6");

                        if (contact.GetAttributeValue<string>("jobtitle") == null || contact.GetAttributeValue<string>("jobtitle") == "")
                        {
                            tracingService.Trace("7");

                            hasTitleCredential = false;
                        }
                        tracingService.Trace("8");

                    }

                    if (!hasTitleCredential)
                    {
                        tracingService.Trace("9");

                        throw new InvalidPluginExecutionException("There is at least one attendee without Title/Credential");
                    }
                    tracingService.Trace("10");

                }
                catch (Exception ex)
                {
                    throw new InvalidPluginExecutionException(ex.Message);
                }
            }
        }
    }
}

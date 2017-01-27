using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Activities;
using System.ServiceModel;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using Microsoft.Xrm.Sdk.Client;

namespace KLITools.XRM.SendEmail
{
    public class SendEmail : CodeActivity
    {
        #region Input Properties
        [Input("Email Template")]
        [ReferenceTarget("template")]
        [Default(null)]
        public InArgument<EntityReference> emailTemplate { get; set; }

        [Input("Sender")]
        public InArgument<string> sender { get; set; }

        [Input("Recipients")]
        public InArgument<string> recipients { get; set; }

        [Input("Set Regarding")]
        [Default("False")]
        public InArgument<bool> setRegarding { get; set; }

        [Input("Custom Fields (eg: {!FieldName:Value;!})")]
        public InArgument<string> customFields { get; set; }
        #endregion

        #region Output Properties
        [Output("IsSuccess")]
        [Default("true")]
        public OutArgument<bool> isSuccess { get; set; }

        [Output("Message")]
        public OutArgument<string> message { get; set; }

        [Output("Email Reference")]
        [ReferenceTarget("email")]
        public OutArgument<EntityReference> outEmail { get; set; }
        #endregion

        protected override void Execute(CodeActivityContext executionContext)
        {
            // Create the tracing service
            ITracingService tracingService = executionContext.GetExtension<ITracingService>();

            if (tracingService == null)
            {
                throw new InvalidPluginExecutionException("Failed to retrieve tracing service.");
            }

            tracingService.Trace("Entered SendCustomerTicketNotification.Execute(), Activity Instance Id: {0}, Workflow Instance Id: {1}",
                executionContext.ActivityInstanceId,
                executionContext.WorkflowInstanceId);

            // Create the context
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();

            if (context == null)
            {
                throw new InvalidPluginExecutionException("Failed to retrieve workflow context.");
            }

            tracingService.Trace("SendCustomerTicketNotification.Execute(), Correlation Id: {0}, Initiating User: {1}",
                context.CorrelationId,
                context.InitiatingUserId);

            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            OrganizationServiceContext sContext = new OrganizationServiceContext(service);

            try
            {
                //Guid entityId = Guid.Empty;
                Entity primaryEntity = null;

                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
                {
                    //entityId = ((Entity)context.InputParameters["Target"]).Id;
                    primaryEntity = ((Entity)context.InputParameters["Target"]);
                }
                else if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is EntityReference)
                {
                    //entityId = ((Entity)context.InputParameters["Target"]).Id;
                    primaryEntity = new Entity(((EntityReference)context.InputParameters["Target"]).LogicalName);
                    primaryEntity.Id = ((EntityReference)context.InputParameters["Target"]).Id;
                }
                else if (context.SharedVariables.Contains("PostBusinessEntity"))
                    primaryEntity = ((Entity)context.SharedVariables["PostBusinessEntity"]);
                else if (context.SharedVariables.Contains("PreBusinessEntity"))
                    primaryEntity = ((Entity)context.SharedVariables["PreBusinessEntity"]);

                //else if (context.InputParameters.Contains("EntityId"))
                //    entityId = new Guid(context.InputParameters["EntityId"].ToString()); 

                if (primaryEntity != null)
                {

                    // Get Sender
                    //EntityReference refFrom = CommonMethods.GetEmailSender(sender.Get(executionContext), service);

                    //if (refFrom == null)
                    //    return;

                    // Email Sender
                    Entity[] emailSender = CommonMethods.ParseRecipients(sender.Get(executionContext).ToLower(), primaryEntity, service);  //CommonMethods.ConvertToPartyList(refFrom);

                    if (emailSender == null || emailSender.Count() <= 0)
                        return;

                    tracingService.Trace("Sender Retrieved");

                    // Read Recipients List from workflow input arguments
                    string[] strRecipientsLst = recipients.Get(executionContext).Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);

                    // Email Recipients Collection
                    Entity[] toRecipient = CommonMethods.ParseRecipients(recipients.Get(executionContext).ToLower(), primaryEntity, service);

                    tracingService.Trace("Recipient Retrieved");

                    if (toRecipient != null && toRecipient.Count() > 0)
                    {
                        // Set Custom Fields
                        string strCustomFields = customFields.Get(executionContext) != null ? customFields.Get(executionContext) : string.Empty;

                        if (setRegarding.Get(executionContext) && !CommonMethods.IsActivityEnabled(primaryEntity.LogicalName, service))
                        {
                            setRegarding.Set(executionContext, false);
                        }

                        Guid emailId = CommonMethods.SendEmail(emailSender,
                                                                toRecipient,
                                                                null, null, null,
                                                                new EntityReference(primaryEntity.LogicalName, primaryEntity.Id),
                                                                emailTemplate.Get(executionContext),
                                                                strCustomFields, !setRegarding.Get(executionContext), service);

                        if (emailId != Guid.Empty)
                        {
                            isSuccess.Set(executionContext, true);
                            message.Set(executionContext, "Email sent to recipents");
                            outEmail.Set(executionContext, new EntityReference("email", emailId));
                        }
                        tracingService.Trace("Email Sent");
                    }
                    else
                    {
                        isSuccess.Set(executionContext, false);
                        message.Set(executionContext, "No recipents found to send an email");
                    }
                }
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                isSuccess.Set(executionContext, false);
                message.Set(executionContext, string.Format("{0}{1}{2}", ex.Message, Environment.NewLine, ex.StackTrace));
            }

            tracingService.Trace("Exiting SendCustomerTicketNotification.Execute(), Correlation Id: {0}", context.CorrelationId);
        }
    }
}

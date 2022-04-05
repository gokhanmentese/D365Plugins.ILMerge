using Microsoft.Xrm.Sdk;
using Crm.ExcelPlugins.Model;
using System;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.ServiceModel;

namespace Crm.ExcelPlugins
{
    public class Plugin : IPlugin
    {
        protected class LocalPluginContext
        {
            internal IServiceProvider ServiceProvider
            {
                get;
                private set;
            }

            internal IOrganizationService OrganizationService
            {
                get;

                private set;
            }

            internal IPluginExecutionContext PluginExecutionContext
            {
                get;

                private set;
            }

            internal IOrganizationServiceFactory OrganizationServiceFactory 
            {
                get;

                private set;
            }

            internal ITracingService TracingService
            {
                get;

                private set;
            }

            internal Entity PluginTargetEntity
            {
                get;

                private set;
            }

            internal EntityReference PluginTargetEntityRef // Delete message
            {
                get;

                private set;
            }

            private LocalPluginContext()
            {
            }

            internal LocalPluginContext(IServiceProvider serviceProvider)
            {
                if (serviceProvider == null)
                {
                    throw new ArgumentNullException("serviceProvider");
                }

                this.PluginExecutionContext = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
                this.TracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
                IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                this.OrganizationServiceFactory = factory;
                this.OrganizationService = factory.CreateOrganizationService(this.PluginExecutionContext.UserId);

                if (this.PluginExecutionContext.InputParameters.Contains("Target"))
                {
                    if (this.PluginExecutionContext.MessageName.Equals(PluginMessages.Create))
                    {
                        this.PluginTargetEntity = (Entity)this.PluginExecutionContext.InputParameters["Target"];
                    }
                    else if (this.PluginExecutionContext.MessageName.Equals(PluginMessages.Update))
                    {
                        this.PluginTargetEntity = (Entity)this.PluginExecutionContext.InputParameters["Target"];
                    }
                    else if (this.PluginExecutionContext.MessageName.Equals(PluginMessages.Delete))
                    {
                        this.PluginTargetEntityRef = (EntityReference)this.PluginExecutionContext.InputParameters["Target"];
                    }
                    else if (this.PluginExecutionContext.MessageName.Equals(PluginMessages.Merge))
                    {
                        this.PluginTargetEntityRef = (EntityReference)this.PluginExecutionContext.InputParameters["Target"];
                    }
                }
            }

            internal void Trace(string message)
            {
                if (string.IsNullOrWhiteSpace(message) || this.TracingService == null)
                {
                    return;
                }

                if (this.PluginExecutionContext == null)
                {
                    this.TracingService.Trace(message);
                }
                else
                {
                    this.TracingService.Trace(
                        "{0}, Correlation Id: {1}, Initiating User: {2}",
                        message,
                        this.PluginExecutionContext.CorrelationId,
                        this.PluginExecutionContext.InitiatingUserId);
                }
            }
        }

        private Collection<Tuple<int, string, string, Action<LocalPluginContext>>> registeredEvents;

        protected Collection<Tuple<int, string, string, Action<LocalPluginContext>>> RegisteredEvents
        {
            get
            {
                if (this.registeredEvents == null)
                {
                    this.registeredEvents = new Collection<Tuple<int, string, string, Action<LocalPluginContext>>>();
                }

                return this.registeredEvents;
            }
        }

        protected string ChildClassName
        {
            get;

            private set;
        }

        public Plugin(Type childClassName)
        {
            this.ChildClassName = childClassName.ToString();
        }

        public void Execute(IServiceProvider serviceProvider)
        {
            if (serviceProvider == null)
            {
                throw new ArgumentNullException("serviceProvider");
            }

            LocalPluginContext localcontext = new LocalPluginContext(serviceProvider);
            localcontext.Trace(string.Format(CultureInfo.InvariantCulture, "Entered {0}.Execute()", this.ChildClassName));

            try
            {
                
                Action<LocalPluginContext> entityAction =
                    (from a in this.RegisteredEvents
                     where (
                     a.Item1 == localcontext.PluginExecutionContext.Stage &&
                     a.Item2 == localcontext.PluginExecutionContext.MessageName &&
                     (string.IsNullOrWhiteSpace(a.Item3) ? true : a.Item3 == localcontext.PluginExecutionContext.PrimaryEntityName)
                     )
                     select a.Item4).FirstOrDefault();

                if (entityAction != null)
                {
                    localcontext.Trace(string.Format(
                        CultureInfo.InvariantCulture,
                        "{0} is firing for Entity: {1}, Message: {2}",
                        this.ChildClassName,
                        localcontext.PluginExecutionContext.PrimaryEntityName,
                        localcontext.PluginExecutionContext.MessageName));

                    entityAction.Invoke(localcontext);

                    return;
                }
            }
            catch (FaultException<OrganizationServiceFault> e)
            {
                localcontext.Trace(string.Format(CultureInfo.InvariantCulture, "Exception: {0}", e.ToString()));
                throw;
            }
            finally
            {
                localcontext.Trace(string.Format(CultureInfo.InvariantCulture, "Exiting {0}.Execute()", this.ChildClassName));
            }
        }
    }
}

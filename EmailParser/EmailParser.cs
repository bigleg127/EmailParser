using System;
using System.Net;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace EmailParser
{
    public class EmailParserPlugin: IPlugin
    {
        /// <summary>
        /// Descriptiopn:   This plugin is designed to read in emails that are stored in Dynamics CRM and asynchronously parse the text into useful fields.
        /// Author:         Darren Hickey
        /// Date:           28/08/2015
        /// </summary>
        public void Execute(IServiceProvider serviceProvider)
        {
            var tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            var serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            var orgservice = serviceFactory.CreateOrganizationService(context.UserId);

            if (!context.InputParameters.Contains("Target") || !(context.InputParameters["Target"] is Entity)) return;
            var email = (Entity)context.InputParameters["Target"];

            if (tracingService == null) return;
            if (email.LogicalName != "email") return;

            var definitions = GetDefinitions(email, orgservice);
            tracingService.Trace(@"Parsing Definitions Loaded\n");

            var enumerable = definitions as IList<Entity> ?? definitions.ToList();
            if (!enumerable.Any())
            {
                tracingService.Trace(@"No Matching Definitions\n");
            }
            else
            {
                foreach (var d in enumerable)
                {
                    var mappings = getMappings(d, orgservice);
                    var body = WebUtility.HtmlDecode(GetEmailBody((string) email.Attributes["description"]));
                    var entityname = (string) d.Attributes["new_crmentityname"];
                    tracingService.Trace(@"Mappings Loaded\n");

                    ParseEmail(body, mappings, entityname, orgservice);
                    tracingService.Trace("Email Parsed\n");
                }
            }
        }

        private static string GetEmailBody(string email)
        {
            const string regex = "^.*<body>";
            const string endregex = "</body>.*$";
            var output = Regex.Replace(email, regex, "", RegexOptions.Singleline);
            output = Regex.Replace(output, endregex, "", RegexOptions.Singleline);
            return output;
        }

        private static void ParseEmail(string emailBody, EntityCollection mappings, string entityType, IOrganizationService service)
        {
            var ent = new Entity(entityType);
            foreach(var e in mappings.Entities)
            {
                var target = (string) e.Attributes["new_emailstring"];
                var index = emailBody.IndexOf(target, StringComparison.Ordinal);
                if (index == -1) continue;
                var startpos = index + target.Length;
                var endpos = emailBody.IndexOf((string)e.Attributes["new_emailstringdelimiter"], startpos, StringComparison.Ordinal);
                var length = endpos - startpos ;
                var result = emailBody.Substring(startpos, length).Trim();
                ent.Attributes.Add((string)e.Attributes["new_crmfieldname"], result);
            }
            service.Create(ent);
        }

        private static IEnumerable<Entity> GetDefinitions(Entity email, IOrganizationService service)
        {
            var subject = (string)email.Attributes["subject"];
            if (subject == null) throw new ArgumentNullException("email");

            var epmappings = new QueryExpression("new_emailparserdefinition");
            var cols = new ColumnSet("new_name", "new_emailparserdefinitionid", "new_crmentityname", "new_subjectpattern", "new_senderaddress");
            epmappings.ColumnSet = cols;
            epmappings.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, "Active"));
            IEnumerable<Entity> result = service.RetrieveMultiple(epmappings).Entities;

            return result
                .Where(element => Regex.IsMatch(subject, (string)element.Attributes["new_subjectpattern"]))
                    .Select(element => element);             
        }

        private EntityCollection getMappings(Entity def, IOrganizationService service)
        {
                var epmappings = new QueryExpression("new_emailparsermapping");
                var cols = new ColumnSet("new_emailstring", "new_crmfieldname", "new_emailstringdelimiter");
                epmappings.ColumnSet = cols;
                epmappings.Criteria.AddCondition( new ConditionExpression("new_definition", 0, def["new_emailparserdefinitionid"]));
                var result = service.RetrieveMultiple(epmappings);
                return result;
        }
    }
}

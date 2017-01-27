using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Messages;
using System.Text.RegularExpressions;

namespace KLITools.XRM.SendEmail
{
    internal static class CommonMethods
    {
        #region Email Recipients
        /// <summary>
        /// Get Email Sender
        /// </summary>
        /// <param name="strKey"></param>
        /// <param name="proxyServiceContext"></param>
        /// <returns></returns>
        public static EntityReference GetEmailSender(string strKey, IOrganizationService service)
        {
            OrganizationServiceContext sContext = new OrganizationServiceContext(service);

            if (strKey == null)
                return null;

            return getPartyByEmail(strKey, sContext);
        }

        public static EntityReference GetRecipientByKey(string strKey, IOrganizationService service)
        {
            OrganizationServiceContext sContext = new OrganizationServiceContext(service);

            if (strKey == null)
                return null;

            return getPartyByEmail(strKey, sContext);
        }
        #endregion

        #region Utility Methods
        /// <summary>
        /// Returns Systems User
        /// </summary>
        /// <param name="strKey"></param>
        /// <param name="sContext"></param>
        /// <returns></returns>
        public static EntityReference GetUserByEmail(string strKey, OrganizationServiceContext sContext)
        {
            return getSystemUser(strKey, sContext);
        }

        public static bool Contains(this string source, string toCheck, StringComparison comp)
        {
            return source != null && toCheck != null && source.IndexOf(toCheck, comp) >= 0;
        }

        /// <summary>
        /// Convert CRM Attribute to String
        /// </summary>
        /// <param name="strAttributeName"></param>
        /// <param name="entObject"></param>
        /// <returns></returns>
        public static string ConvertAttributeToString(string strAttributeName, Entity entObject, OrganizationServiceContext sContext)
        {
            string strValue = String.Empty;
            if (entObject[strAttributeName].GetType().ToString() != "Microsoft.Xrm.Sdk.AliasedValue")
            {
                switch (entObject[strAttributeName].GetType().ToString())
                {
                    case "Microsoft.Xrm.Sdk.EntityReference":
                        strValue = ((EntityReference)entObject[strAttributeName]).Name;
                        break;
                    case "Microsoft.Xrm.Sdk.OptionSetValue":
                        strValue = entObject.FormattedValues[strAttributeName];
                        break;
                    case "System.DateTime":
                        strValue = (Convert.ToDateTime(entObject[strAttributeName].ToString())).ToLocalTime().ToString();
                        break;
                    case "System.Boolean":
                        strValue = (bool)entObject[strAttributeName] ? "Yes" : "No";
                        break;
                    case "System.Int32":
                        if (strAttributeName.Contains("timezone"))
                            strValue = getTimeZoneLabel(Convert.ToInt32(entObject[strAttributeName].ToString()), sContext);
                        else
                            strValue = entObject[strAttributeName].ToString();
                        break;
                    default:
                        strValue = entObject[strAttributeName].ToString();
                        break;
                }
            }
            else
            {
                switch (entObject.GetAttributeValue<AliasedValue>(strAttributeName).Value.GetType().ToString())
                {
                    case "Microsoft.Xrm.Sdk.EntityReference":
                        strValue = ((EntityReference)entObject.GetAttributeValue<AliasedValue>(strAttributeName).Value).Name;
                        break;
                    case "Microsoft.Xrm.Sdk.OptionSetValue":
                        //strValue = ((OptionSetValue)entObject.GetAttributeValue<AliasedValue>(strAttributeName).Value).Value.ToString();
                        //strValue = entObject.FormattedValues.Where(e => e.Key == strAttributeName).FirstOrDefault().Value;
                        strValue = entObject.FormattedValues[strAttributeName];
                        break;
                    case "System.DateTime":
                        strValue = (Convert.ToDateTime(entObject.GetAttributeValue<AliasedValue>(strAttributeName).Value)).ToLocalTime().ToString();
                        break;
                    case "System.Boolean":
                        strValue = (bool)entObject.GetAttributeValue<AliasedValue>(strAttributeName).Value ? "Yes" : "No";
                        break;
                    case "System.Int32":
                        if (strAttributeName.Contains("timezone"))
                            strValue = getTimeZoneLabel(Convert.ToInt32(entObject.GetAttributeValue<AliasedValue>(strAttributeName).Value.ToString()), sContext);
                        else
                            strValue = entObject.GetAttributeValue<AliasedValue>(strAttributeName).Value.ToString();
                        break;
                    default:
                        strValue = entObject.GetAttributeValue<AliasedValue>(strAttributeName).Value.ToString();
                        break;
                }
            }
            return strValue;
        }

        /// <summary>
        /// Verify if Activities are enabled for entity
        /// </summary>
        /// <param name="entLogicalName"></param>
        /// <param name="service"></param>
        /// <returns></returns>
        public static bool IsActivityEnabled(string entLogicalName, IOrganizationService service)
        {
            var metaResponse = (RetrieveEntityResponse)service.Execute(new RetrieveEntityRequest
            {
                EntityFilters = EntityFilters.Relationships,
                LogicalName = entLogicalName,
                RetrieveAsIfPublished = true
            });

            return
                metaResponse.EntityMetadata.OneToManyRelationships
                .Any(r => r.ReferencingEntity == "activitypointer");
        }
        #endregion

        #region Activity Party
        /// <summary>
        /// 
        /// </summary>
        /// <param name="entRefColl"></param>
        /// <returns></returns>
        public static Entity[] ConvertToPartyList(List<EntityReference> lstEntityReferences)
        {
            if (lstEntityReferences == null || lstEntityReferences.Count == 0)
                return null;

            Entity[] partyList = new Entity[lstEntityReferences.Count];

            int index = 0;
            foreach (EntityReference entRef in lstEntityReferences)
            {
                partyList[index++] = ConvertToActivityParty(entRef);
            }

            return partyList;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="entRefColl"></param>
        /// <returns></returns>
        public static Entity[] ConvertToPartyList(EntityReferenceCollection entRefColl)
        {
            if (entRefColl == null || entRefColl.Count == 0)
                return null;

            Entity[] partyList = new Entity[entRefColl.Count];

            int index = 0;
            foreach (EntityReference entRef in entRefColl)
            {
                partyList[index++] = ConvertToActivityParty(entRef);
            }

            return partyList;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="entRef"></param>
        /// <returns></returns>
        public static Entity[] ConvertToPartyList(EntityReference entRef)
        {
            Entity[] partyList = new Entity[] { ConvertToActivityParty(entRef) };

            return partyList;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="entRef"></param>
        /// <returns></returns>
        public static Entity ConvertToActivityParty(EntityReference entRef)
        {
            Entity party = new Entity("activityparty");
            party["partyid"] = entRef;

            return party;
        }
        #endregion

        #region Send Email
        public static Guid SendEmail(Entity[] emailSender, Entity[] toRecipient, Entity[] ccRecipients, string strSubject, string strDescription, EntityReference regardingObj, EntityReference templateRef, string strCustomFields, bool ignoreRegarding, IOrganizationService service)
        {
            Microsoft.Crm.Sdk.Messages.InstantiateTemplateRequest instTemplateReq;
            Entity email;
            int subjectMaxLength = 255;

            // Email using template
            if (templateRef != null && templateRef.Id != Guid.Empty)
            {
                instTemplateReq = new Microsoft.Crm.Sdk.Messages.InstantiateTemplateRequest
                {
                    TemplateId = templateRef.Id,
                    ObjectId = regardingObj.Id,
                    ObjectType = regardingObj.LogicalName
                };

                Microsoft.Crm.Sdk.Messages.InstantiateTemplateResponse instTemplateResp = (Microsoft.Crm.Sdk.Messages.InstantiateTemplateResponse)service.Execute(instTemplateReq);

                email = (Entity)instTemplateResp.EntityCollection.Entities[0];

                if (strCustomFields != string.Empty)
                {
                    Regex regEx = new Regex(@"(?<=\{!).*?(?=;!\})", RegexOptions.Singleline);
                    var matches = regEx.Matches(strCustomFields);

                    foreach (Match match in matches)
                    {
                        var customField = match.ToString().Split(':');
                        if (customField.Length > 0)
                        {
                            string strValue = Regex.Replace(match.ToString().Replace(customField[0] + ":", ""), @"\r\n?|\n", "<br>");

                            email["description"] = email.GetAttributeValue<string>("description").Replace("{!Custom: " + customField[0] + ";}", strValue).Replace("{!Custom:" + customField[0] + ";}", strValue);

                            email["subject"] = email.GetAttributeValue<string>("subject").Replace("{!Custom: " + customField[0] + ";}", strValue).Replace("{!Custom:" + customField[0] + ";}", strValue);
                        }
                    }
                }

                // Parse EMail Subject
                email["subject"] = ParseEmailBody(email.GetAttributeValue<string>("subject"), regardingObj, service);

                // Parse Email Body
                email["description"] = ParseEmailBody(email.GetAttributeValue<string>("description"), regardingObj, service);
            }
            else // Email without using template
            {
                email = new Entity("email");
                if (!String.IsNullOrEmpty(strDescription))
                    email["description"] = strDescription;
            }

            email["from"] = emailSender;
            email["to"] = toRecipient;
            email["cc"] = ccRecipients;

            if (!String.IsNullOrEmpty(strSubject))
                email["subject"] = strSubject;

            if (email.GetAttributeValue<string>("subject").Length > subjectMaxLength)
                email["subject"] = email.GetAttributeValue<string>("subject").Substring(0, subjectMaxLength - 3) + "...";

            email["regardingobjectid"] = ignoreRegarding == true ? null : regardingObj;

            Guid emailId = service.Create(email);

            if (emailId != Guid.Empty)
            {
                Microsoft.Crm.Sdk.Messages.SendEmailRequest sendEmailreq = new Microsoft.Crm.Sdk.Messages.SendEmailRequest
                {
                    EmailId = emailId,
                    TrackingToken = null,
                    IssueSend = true
                };

                Microsoft.Crm.Sdk.Messages.SendEmailResponse sendEmailresp = (Microsoft.Crm.Sdk.Messages.SendEmailResponse)service.Execute(sendEmailreq);
            }

            return emailId;
        }
        #endregion

        #region Private Methods
        /// <summary>
        /// 
        /// </summary>
        /// <param name="strEmail"></param>
        /// <param name="sContext"></param>
        /// <returns></returns>
        private static EntityReference getPartyByEmail(string strEmail, OrganizationServiceContext sContext)
        {
            return getPartyByEmail(strEmail, string.Empty, sContext);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="strEmail"></param>
        /// <param name="strEntityLogicalName"></param>
        /// <param name="proxyServiceContext"></param>
        /// <returns></returns>
        private static EntityReference getPartyByEmail(string strEmail, string strEntityLogicalName, OrganizationServiceContext sContext)
        {
            EntityReference entRef = null;

            // Check if it is Queue
            if (strEntityLogicalName == string.Empty || strEntityLogicalName.ToLower() == "queue")
                entRef = getQueue(strEmail, sContext);

            if (entRef == null && (strEntityLogicalName == string.Empty || strEntityLogicalName.ToLower() == "systemuser"))
            {
                // Check if it is User
                entRef = getSystemUser(strEmail, sContext);
            }

            if (entRef == null && (strEntityLogicalName == string.Empty || strEntityLogicalName.ToLower() == "systemuser"))
            {
                // Check if it is Contact
                entRef = getContactByEmail(strEmail, sContext);

                if (entRef == null)
                    entRef = getContactByName(strEmail, sContext);
            }

            return entRef;
        }

        /// <summary>
        /// Returns Queue
        /// </summary>
        /// <param name="strKey"></param>
        /// <param name="sContext"></param>
        /// <returns></returns>
        private static EntityReference getQueue(string strKey, OrganizationServiceContext sContext)
        {
            EntityReference entRef = null;

            var entQueue = (from q in sContext.CreateQuery("queue")
                            where q.GetAttributeValue<string>("emailaddress") == strKey || q.GetAttributeValue<string>("name") == strKey
                            select new
                            {
                                Id = q.Id,
                                Name = q.GetAttributeValue<string>("name")
                            }).FirstOrDefault();

            if (entQueue != null)
                entRef = new EntityReference("queue", entQueue.Id);

            return entRef;
        }

        /// <summary>
        /// Returns System User
        /// </summary>
        /// <param name="strKey"></param>
        /// <param name="sContext"></param>
        /// <returns></returns>
        private static EntityReference getSystemUser(string strKey, OrganizationServiceContext sContext)
        {
            EntityReference entRef = null;

            var entUser = (from u in sContext.CreateQuery("systemuser")
                           where u.GetAttributeValue<string>("internalemailaddress") == strKey || u.GetAttributeValue<string>("personalemailaddress") == strKey || u.GetAttributeValue<string>("domainname").Contains(strKey) || u.GetAttributeValue<string>("fullname").Contains(strKey)
                           select new
                           {
                               Id = u.Id,
                               FirstName = u.GetAttributeValue<string>("firstname"),
                               LastName = u.GetAttributeValue<string>("lastname")
                           }).FirstOrDefault();

            if (entUser != null)
                entRef = new EntityReference("systemuser", entUser.Id);

            return entRef;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="strEmail"></param>
        /// <param name="sContext"></param>
        /// <returns></returns>
        private static EntityReference getContactByEmail(string strEmail, OrganizationServiceContext sContext)
        {
            EntityReference entRef = null;
            // check if it is Contact
            var entContact = (from c in sContext.CreateQuery("contact")
                              where c.GetAttributeValue<string>("emailaddress1") == strEmail || c.GetAttributeValue<string>("emailaddress2") == strEmail || c.GetAttributeValue<string>("emailaddress3") == strEmail
                              select new
                              {
                                  Id = c.Id,
                                  FirstName = c.GetAttributeValue<string>("firstname"),
                                  LastName = c.GetAttributeValue<string>("lastname")
                              }).FirstOrDefault();

            if (entContact != null)
                entRef = new EntityReference("contact", entContact.Id);

            return entRef;
        }

        /// <summary>
        /// Get Contact Id using Contact Fullname
        /// </summary>
        /// <param name="strFullName"></param>
        /// <param name="service"></param>
        /// <returns></returns>
        private static EntityReference getContactByName(string strFullName, OrganizationServiceContext sContext)
        {
            EntityReference entRef = null;

            var entContact = (from c in sContext.CreateQuery("contact")
                              where c.GetAttributeValue<string>("fullname") == strFullName && c.GetAttributeValue<OptionSetValue>("statecode").Value == 0
                              select new
                              {
                                  Id = c.Id,
                                  FirstName = c.GetAttributeValue<string>("firstname"),
                                  LastName = c.GetAttributeValue<string>("lastname")
                              }).FirstOrDefault();

            if (entContact != null && entContact.Id != null)
                entRef = new EntityReference("contact", entContact.Id);

            return entRef;

        }

        /// <summary>
        /// Parse Email Body
        /// </summary>
        /// <param name="strBody"></param>
        /// <param name="entRefRegarding"></param>
        /// <param name="organizationService"></param>
        /// <returns></returns>
        private static string ParseEmailBody(string strBody, EntityReference entRefRegarding, IOrganizationService organizationService)
        {
            OrganizationServiceContext sContext = new OrganizationServiceContext(organizationService);

            List<PlaceHolder> lstPlaceHolders = getPlaceHolders(strBody);

            if (lstPlaceHolders.Count <= 0)
                return strBody;

            EntityCollection entColl = organizationService.RetrieveMultiple(new Microsoft.Xrm.Sdk.Query.FetchExpression(
                getFetchXMLFromPlaceHolders(entRefRegarding.LogicalName.ToLower(), entRefRegarding.Id, lstPlaceHolders)));

            if (entColl != null && entColl.Entities.Count > 0)
            {
                Entity entRegarding = entColl.Entities[0];
                foreach (PlaceHolder p in lstPlaceHolders)
                {
                    // Set Replace Value for Placeholder
                    //lstPlaceHolders.Where(o => o.Attribute == p.Attribute).ToList().ForEach(o => o.Value = entComment[p.Attribute].ToString());
                    if (entRegarding.Contains(p.SourceAttribute) && entRegarding[p.SourceAttribute] != null)
                    {
                        if (entRegarding[p.SourceAttribute].GetType().ToString() == "Microsoft.Xrm.Sdk.EntityReference")
                        {
                            EntityReference entRefSourceAttribute = ((EntityReference)entRegarding[p.SourceAttribute]);

                            EntityCollection relatedEntityColl = organizationService.RetrieveMultiple(new Microsoft.Xrm.Sdk.Query.FetchExpression(
                                getFetchXML(entRefSourceAttribute.LogicalName.ToLower(), entRefSourceAttribute.Id, p.RelatedAttribute.ToLower())));

                            if (relatedEntityColl != null && relatedEntityColl.Entities.Count > 0)
                            {
                                Entity entRelated = relatedEntityColl.Entities[0];

                                if (entRelated.Contains(p.RelatedAttribute) && entRelated[p.RelatedAttribute] != null)
                                {
                                    string strValue = ConvertAttributeToString(p.RelatedAttribute, entRelated, sContext);

                                    strBody = Regex.Replace(strBody, p.Source, strValue, RegexOptions.IgnoreCase);//strBody.Replace(p.Source, strValue);
                                }
                                else
                                {
                                    strBody = Regex.Replace(strBody, p.Source, "- NA -", RegexOptions.IgnoreCase);//strBody.Replace(p.Source, "- NA -");
                                }
                            }
                        }
                    }
                    else
                    {
                        strBody = strBody.Replace(p.Source, "- NA -");
                    }
                }
            }

            return strBody;
        }

        /// <summary>
        /// Parse Recipients from Input String and return Party List
        /// </summary>
        /// <param name="strRecipients"></param>
        /// <param name="primaryEntity"></param>
        /// <param name="organizationService"></param>
        /// <returns></returns>
        public static Entity[] ParseRecipients(string strRecipients, Entity primaryEntity, IOrganizationService organizationService)
        {
            List<EntityReference> lstRecipients = new List<EntityReference>();

            List<PlaceHolder> lstPlaceHolders = getPlaceHolders(strRecipients);

            EntityCollection entColl = organizationService.RetrieveMultiple(new Microsoft.Xrm.Sdk.Query.FetchExpression(
                 getFetchXMLFromPlaceHolders(primaryEntity.LogicalName, primaryEntity.Id, lstPlaceHolders)));

            if (entColl != null && entColl.Entities.Count > 0)
            {
                Entity entPrimaryEntity = entColl.Entities[0];
                foreach (PlaceHolder p in lstPlaceHolders)
                {
                    // Set Replace Value for Placeholder
                    //lstPlaceHolders.Where(o => o.Attribute == p.Attribute).ToList().ForEach(o => o.Value = entComment[p.Attribute].ToString());
                    if (p.RelatedAttribute == null && entPrimaryEntity.Contains(p.SourceAttribute) && entPrimaryEntity[p.SourceAttribute] != null &&
                        entPrimaryEntity[p.SourceAttribute].GetType().ToString() == "Microsoft.Xrm.Sdk.EntityReference")
                    {
                        lstRecipients.Add(((EntityReference)entPrimaryEntity[p.SourceAttribute]));
                    }
                    else if (p.RelatedAttribute != null && entPrimaryEntity.Contains(p.SourceAttribute) && entPrimaryEntity[p.SourceAttribute] != null &&
                        entPrimaryEntity[p.SourceAttribute].GetType().ToString() == "Microsoft.Xrm.Sdk.EntityReference")
                    {
                        EntityReference entRefSourceAttribute = ((EntityReference)entPrimaryEntity[p.SourceAttribute]);

                        EntityCollection relatedEntityColl = organizationService.RetrieveMultiple(
                            new Microsoft.Xrm.Sdk.Query.FetchExpression(getFetchXML(entRefSourceAttribute.LogicalName.ToLower(), entRefSourceAttribute.Id, p.RelatedAttribute.ToLower())));
                        if (relatedEntityColl != null && relatedEntityColl.Entities.Count > 0)
                        {
                            Entity entRelated = relatedEntityColl.Entities[0];

                            if (entRelated.Contains(p.RelatedAttribute) && entRelated[p.RelatedAttribute] != null && entRelated[p.RelatedAttribute].GetType().ToString() == "Microsoft.Xrm.Sdk.EntityReference")
                            {
                                lstRecipients.Add(((EntityReference)entRelated[p.RelatedAttribute]));
                            }
                        }
                    }
                }
            }

            // Read Recipients List from workflow input arguments
            string[] strRecipientsLst = strRecipients.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);

            foreach (string strRep in strRecipientsLst)
            {
                EntityReference entRep = GetRecipientByKey(strRep, organizationService);
                if (entRep != null)
                    lstRecipients.Add(entRep);
            }

            if (lstRecipients.Count > 0)
                return ConvertToPartyList(lstRecipients);
            else
                return null;
        }

        /// <summary>
        /// Parse & Convert the String to Placeholder object
        /// </summary>
        /// <param name="strSourceString"></param>
        /// <param name="strKey"></param>
        /// <returns></returns>
        private static List<PlaceHolder> getPlaceHolders(string strSourceString, string strKey = @"(?<=\{!).*?(?=;\})")
        {
            List<PlaceHolder> lstPlaceHolders = new List<PlaceHolder>();

            // Parse placeholders from the string
            Regex regEx = new Regex(strKey);
            var matches = regEx.Matches(strSourceString);

            foreach (Match match in matches)
            {
                PlaceHolder placeHolder = new PlaceHolder();

                placeHolder.Source = match.ToString().ToLower();
                placeHolder.Source = "{!" + (Regex.Replace(Regex.Replace(placeHolder.Source, @"\{!", String.Empty), @";\}", String.Empty)) + ";}";

                var strParse = (Regex.Replace(Regex.Replace(Regex.Replace(placeHolder.Source, @"\s+", String.Empty), @"\{!", String.Empty), @";\}", String.Empty)).Split(':');
                placeHolder.SourceEntity = strParse[0];
                placeHolder.SourceAttribute = strParse[1];

                if (placeHolder.Source.Contains("{!related!", StringComparison.OrdinalIgnoreCase))
                {
                    strParse = (Regex.Replace(Regex.Replace(Regex.Replace(placeHolder.Source, @"\s+", String.Empty), @"\{!related!", String.Empty, RegexOptions.IgnoreCase), @";\}", String.Empty)).Split(':');
                    placeHolder.SourceEntity = "";
                    placeHolder.SourceAttribute = strParse[0];
                    placeHolder.RelatedAttribute = strParse[1];
                }

                lstPlaceHolders.Add(placeHolder);
            }

            return lstPlaceHolders;
        }

        /// <summary>
        /// Build FetchXML using PlaceHolder objects Source Attributies
        /// </summary>
        /// <param name="entityName"></param>
        /// <param name="entityId"></param>
        /// <param name="lstPlaceHolders"></param>
        /// <returns></returns>
        private static string getFetchXMLFromPlaceHolders(string entityName, Guid entityId, List<PlaceHolder> lstPlaceHolders)
        {
            // Fetch Query
            var fetchString = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                                  <entity name='{0}'>
                                                    {1}
                                                     <filter type='and'>
                                                         <condition attribute='{0}id' operator='eq' value='{2}' />
                                                     </filter>
                                                  </entity>
                                                </fetch>";

            string fetchXMLAttributes = string.Empty;

            foreach (PlaceHolder p in lstPlaceHolders)
            {
                if (p.SourceAttribute != null && !fetchXMLAttributes.Contains(p.SourceAttribute))
                    fetchXMLAttributes += string.Format("<attribute name='{0}' />", p.SourceAttribute);
            }

            return string.Format(fetchString, entityName, fetchXMLAttributes, entityId);
        }

        /// <summary>
        /// Build FetchXML using Attribute Name
        /// </summary>
        /// <param name="entityName"></param>
        /// <param name="entityId"></param>
        /// <param name="attributeName"></param>
        /// <returns></returns>
        private static string getFetchXML(string entityName, Guid entityId, string attributeName)
        {
            // Retrieve 
            string fetchString = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                               <entity name='{0}'>
                                                <attribute name='{1}' />
                                                 <filter type='and'>
                                                     <condition attribute='{2}' operator='eq' value='{3}' />
                                                 </filter>
                                            </entity>
                                            </fetch>";

            return String.Format(fetchString, entityName.ToLower(), attributeName.ToLower(), entityName.ToLower() + "id", entityId);
        }

        private static string getTimeZoneLabel(int intTimeZoneCode, OrganizationServiceContext sContext)
        {
            string strTimeZone = string.Empty;

            var entTimeZone = (from t in sContext.CreateQuery("timezonedefinition")
                               where t.GetAttributeValue<int>("timezonecode") == intTimeZoneCode
                               select new
                               {
                                   StandardName = t.GetAttributeValue<string>("standardname")
                               }).FirstOrDefault();

            if (entTimeZone != null && entTimeZone.StandardName != null)
                strTimeZone = entTimeZone.StandardName;

            return strTimeZone;
        }
        #endregion
    }

    class PlaceHolder
    {
        public string Source { get; set; }
        public string SourceEntity { get; set; }
        public string SourceAttribute { get; set; }
        public string RelatedEntity { get; set; }
        public string RelatedAttribute { get; set; }
        public string Value { get; set; }
    }
}

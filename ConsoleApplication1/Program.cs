using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Web.Script.Serialization;
using System.Xml.Linq;
using System.Linq;
using System.Security;
using System.Text;
using System.Xml;

namespace ConsoleApplication1
{
    class Program
    {
        static void Main(string[] args)
        {
            System.Security.SecureString password = new System.Security.SecureString();
            foreach (char c in "SsHh1620".ToCharArray())
                password.AppendChar(c);
            SharePointOnlineCredentials credentials = new SharePointOnlineCredentials("sharepoint@hiberus-sp.com", password);
            //credenciales del usuario introducido
            using (ClientContext clientContext = new ClientContext("https://hiberussp.sharepoint.com/sites/pnuin"))
            {
                clientContext.Credentials = credentials;

                //Web web = clientContext.Web;
                //clientContext.Load(web);
                //clientContext.ExecuteQuery();

                //List l = clientContext.Web.Lists.GetByTitle("FAQs");
                //clientContext.Load(l);
                //clientContext.ExecuteQuery();

                //ViewCollection views = l.Views;
                //clientContext.Load(views);
                //clientContext.ExecuteQuery();

                ////View v = l.Views.GetByTitle("FAQs");
                ////clientContext.Load(v);
                ////clientContext.ExecuteQuery();
                ////v.JSLink = "~site/_catalogs/masterpage/Adjuntos.js";
                ////v.Update();
                ////l.Update();
                ////web.Update();
                ////clientContext.Load(v);
                ////clientContext.ExecuteQuery();

                //foreach (var v in l.Views)
                //{
                //    if (v.ServerRelativeUrl.Contains("FAQ"))
                //    {
                //        Microsoft.SharePoint.Client.File file = web.GetFileByServerRelativeUrl(v.ServerRelativeUrl);
                //        LimitedWebPartManager wpm = file.GetLimitedWebPartManager(PersonalizationScope.Shared);
                //        web.Context.Load(wpm.WebParts, wps => wps.Include(wp => wp.WebPart.Title));
                //        web.Context.ExecuteQuery();

                //        //Set the properties for all web parts
                //        foreach (WebPartDefinition wpd in wpm.WebParts)
                //        {
                //            WebPart wp = wpd.WebPart;
                //            wp.Properties["JSLink"] = "~site/_catalogs/masterpage/Adjuntos.js";
                //            wpd.SaveWebPartChanges();
                //            web.Context.ExecuteQuery();
                //        }
                //    }
                //}

                //ListCreationInformation listCreationInformation = new ListCreationInformation();
                //listCreationInformation.Title = "Encuesta6";
                //listCreationInformation.TemplateType = 102;
                //listCreationInformation.TemplateFeatureId = new Guid("00bfea71-eb8a-40b1-80c7-506be7590102");
                //listCreationInformation.Description = "Encuesta 5";
                //listCreationInformation.QuickLaunchOption = QuickLaunchOptions.Off;
                //XmlDocument doc = new XmlDocument();
                //doc.Load("../../Survey.xml");
                //listCreationInformation.CustomSchemaXml = doc.OuterXml;//variable to store schemaXML as string reading schema.xml file
                //clientContext.Web.Lists.Add(listCreationInformation);
                //clientContext.ExecuteQuery();

                string digest = GetFormDigest("https://hiberussp.sharepoint.com/sites/pnuin",credentials);
                //UpdateList(digest, "https://hiberussp.sharepoint.com/sites/pnuin", credentials);
                //getLists(digest, "https://hiberussp.sharepoint.com/sites/pnuin", credentials);
                //UpdateList2(digest, "https://hiberussp.sharepoint.com/sites/pnuin", credentials);
                UpdateList3("https://hiberussp.sharepoint.com/sites/pnuin", credentials);
            }

        }
       

        private static string GetFormDigest(string webUrl, ICredentials credentials)
        {
            //Validate input
            if (String.IsNullOrEmpty(webUrl) || String.IsNullOrWhiteSpace(webUrl))
                return String.Empty;

            //Create REST Request
            Uri uri = new Uri(webUrl + "/_api/contextinfo");
            HttpWebRequest restRequest = (HttpWebRequest)WebRequest.Create(uri);
            restRequest.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED:f");
            restRequest.Credentials = credentials;
            restRequest.Method = "POST";
            restRequest.ContentLength = 0;

            //Retrieve Response
            HttpWebResponse restResponse = (HttpWebResponse)restRequest.GetResponse();
            XDocument atomDoc = XDocument.Load(restResponse.GetResponseStream());
            XNamespace d = "http://schemas.microsoft.com/ado/2007/08/dataservices";

            //Extract Form Digest
            return atomDoc.Descendants(d + "FormDigestValue").First().Value;
        }

        public static async void UpdateList(string digest,string webUrl,ICredentials credentials)
        {

            Uri uri = new Uri(webUrl + "/_api/web/lists/GetByTitle('Encuesta1')");
            //start replacement
            HttpClientHandler httpClientHandler = new HttpClientHandler();
            httpClientHandler.Credentials = credentials;
            HttpClient client = new HttpClient(httpClientHandler);
            client.DefaultRequestHeaders.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
            client.DefaultRequestHeaders.Add("Accept", "application/json;odata=verbose");
            client.DefaultRequestHeaders.Add("ContentType", "application/json;odata=verbose");
            client.DefaultRequestHeaders.Add("X-RequestDigest", digest);
            client.DefaultRequestHeaders.Add("X-HTTP-Method", "Merge");
            client.DefaultRequestHeaders.Add("IF-MATCH", "*");

            HttpContent content = new StringContent("{ '__metadata': { 'type': 'SP.List' }, 'Title': 'Encuesta3' }");
            HttpResponseMessage response = await client.PostAsync(uri, content);
            response.EnsureSuccessStatusCode();
        }

        public static void UpdateList2(string digest, string webUrl, ICredentials credentials)
        {
            HttpWebRequest endpointRequest = (HttpWebRequest)HttpWebRequest.Create(webUrl + "/_api/web/lists/GetByTitle('Encuesta1')");
            endpointRequest.Method = "POST"; 
            endpointRequest.Credentials = credentials;
            endpointRequest.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
            endpointRequest.Accept = "application/json;odata=verbose";
            endpointRequest.ContentType = "application/json;odata=verbose";
            endpointRequest.Headers.Add("X-RequestDigest", digest);
            endpointRequest.Headers.Add("IF-MATCH", "*");
            endpointRequest.Headers["X-HTTP-Method"] = "MERGE";

            string stringData = "{ '__metadata': { 'type': 'SP.List' }, 'ReadSecurity': 2 , 'WriteSecurity' : 2 , 'NoCrawl' : true}";
            //string stringData = "{ '__metadata': { 'type': 'SP.List' }, 'AllowMultiResponses' : true }";
            
            endpointRequest.ContentLength = stringData.Length;
            StreamWriter writer = new StreamWriter(endpointRequest.GetRequestStream());
            writer.Write(stringData);
            writer.Flush();

            HttpWebResponse endpointResponse = (HttpWebResponse)endpointRequest.GetResponse();
        }

        public static void getLists(string digest,string webUrl,ICredentials credentials)
        {
            HttpWebRequest endpointRequest = (HttpWebRequest)HttpWebRequest.Create(webUrl + "/_api/web/lists");
            endpointRequest.Method = "GET";
            endpointRequest.Credentials = credentials;
            endpointRequest.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
            endpointRequest.Accept = "application/json;odata=verbose";
            endpointRequest.ContentType = "application/json;odata=verbose";
            endpointRequest.Headers.Add("X-RequestDigest", digest);
            //endpointRequest.Headers.Add("Authorization", "Bearer " + accessToken);
            //HttpWebResponse endpointResponse = (HttpWebResponse)endpointRequest.GetResponse();
            string x = String.Empty;
            using (var response = (HttpWebResponse)endpointRequest.GetResponse())
            {
                var encoding = Encoding.GetEncoding(response.CharacterSet);
                
                using (var responseStream = response.GetResponseStream())
                {
                    using (var reader = new StreamReader(responseStream, encoding))
                    {
                        x = reader.ReadToEnd();
                    }
                }
            }
            string kk = "";
        }


        public static void UpdateList3(string webUrl,SharePointOnlineCredentials credentials)
        {
            List_Service.Lists listWS = new List_Service.Lists();
            //listWS.CookieContainer = GetO365CookieContainer(credentials,webUrl);
            listWS.CookieContainer = new CookieContainer();
            listWS.CookieContainer.Add(GetFedAuthCookie(credentials, webUrl));
            //listWS.Credentials = credentials;
            listWS.UseDefaultCredentials = false;

            //XmlNode ndList = listWS.GetListCollection();
            //string kk = string.Empty;

            XmlNode ndList = listWS.GetList("Encuesta1");
            XmlNode ndVersion = ndList.Attributes["Version"];
            //XmlNode ndID = ndList.Attributes["ID"];
            XmlDocument xmlDoc = new System.Xml.XmlDocument();

            XmlNode ndProperties = xmlDoc.CreateNode(XmlNodeType.Element, "List","");
            XmlAttribute ndAllowMultiResponses = (XmlAttribute)xmlDoc.CreateNode(XmlNodeType.Attribute,"AllowMultiResponses","");

            ndAllowMultiResponses.Value = "TRUE";
            ndProperties.Attributes.Append(ndAllowMultiResponses);

            XmlNode ndReturn = listWS.UpdateList("Encuesta1", ndProperties, null, null, null,ndVersion.Value);

        }

        //public static CookieContainer GetAuthCookies(Uri webUri,SharePointOnlineCredentials credentials)
        //{
        //    var authCookie = credentials.GetAuthenticationCookie(webUri);
        //    var cookieContainer = new CookieContainer();
        //    cookieContainer.SetCookies(webUri, authCookie);
        //    return cookieContainer;
        //}

        public static CookieContainer GetO365CookieContainer(SharePointOnlineCredentials credentials, string targetSiteUrl)
        {

            Uri targetSite = new Uri(targetSiteUrl);
            string cookieString = credentials.GetAuthenticationCookie(targetSite);
            CookieContainer container = new CookieContainer();
            string trimmedCookie = cookieString.TrimStart("SPOIDCRL=".ToCharArray());
            container.Add(new Cookie("FedAuth", trimmedCookie, string.Empty, targetSite.Authority));
            return container;
        }

        private static Cookie GetFedAuthCookie(SharePointOnlineCredentials credentials,string webUrl)
        {
            string authCookie = credentials.GetAuthenticationCookie(new Uri(webUrl));
            if (authCookie.Length > 0)
            {
                return new Cookie("SPOIDCRL", authCookie.TrimStart("SPOIDCRL=".ToCharArray()), String.Empty, new Uri(webUrl).Authority);
            }
            else
            {
                return null;
            }
        }



    }

   
}

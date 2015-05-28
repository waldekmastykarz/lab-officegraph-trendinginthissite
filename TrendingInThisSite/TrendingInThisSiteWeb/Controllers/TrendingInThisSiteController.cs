using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Search.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;
using TrendingInThisSiteWeb.Models;
using TrendingInThisSiteWeb.Extensions;

namespace TrendingInThisSiteWeb.Controllers {
    public class TrendingInThisSiteController : Controller {
        [SharePointContextFilter]
        public ActionResult Index() {
            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);

            IEnumerable<TrendingDocument> trendingDocuments = null;

            using (var clientContext = spContext.CreateUserClientContextForSPHost()) {
                if (clientContext != null) {
                    string[] siteMembersEmails = GetSiteMembersEmails(spContext);
                    string[] actorIds = GetActorIds(clientContext, siteMembersEmails);
                    trendingDocuments = GetTrendingDocuments(clientContext, actorIds);
                }
            }

            return View(trendingDocuments);
        }

        private static string[] GetSiteMembersEmails(SharePointContext spContext) {
            List<string> siteMembersEmails = null;

            using (var clientContext = spContext.CreateAppOnlyClientContextForSPHost()) {
                Web web = clientContext.Web;
                clientContext.Load(web, w => w.AssociatedMemberGroup.Users.Include(u => u.Email));
                clientContext.ExecuteQuery();

                siteMembersEmails = new List<string>(web.AssociatedMemberGroup.Users.Count);

                foreach (User user in web.AssociatedMemberGroup.Users) {
                    if (!String.IsNullOrEmpty(user.Email)) {
                        siteMembersEmails.Add(user.Email);
                    }
                }
            }

            return siteMembersEmails.ToArray();
        }

        private static string[] GetActorIds(ClientContext clientContext, string[] siteMembersEmails) {
            StringBuilder searchQueryText = new StringBuilder();

            foreach (string userEmail in siteMembersEmails) {
                if (searchQueryText.Length > 0) {
                    searchQueryText.Append(" OR ");
                }

                searchQueryText.AppendFormat("UserName:{0}", userEmail);
            }

            KeywordQuery searchQuery = new KeywordQuery(clientContext);
            searchQuery.QueryText = searchQueryText.ToString();
            searchQuery.SelectProperties.Add("DocId");
            searchQuery.SourceId = new Guid("b09a7990-05ea-4af9-81ef-edfab16c4e31");
            searchQuery.RowLimit = 100;

            SearchExecutor searchExecutor = new SearchExecutor(clientContext);
            ClientResult<ResultTableCollection> results = searchExecutor.ExecuteQuery(searchQuery);
            clientContext.ExecuteQuery();

            List<string> actorIds = new List<string>(results.Value[0].ResultRows.Count());

            foreach (var row in results.Value[0].ResultRows) {
                actorIds.Add(row["DocId"].ToString());
            }

            return actorIds.ToArray();
        }

        private static IEnumerable<TrendingDocument> GetTrendingDocuments(ClientContext clientContext, string[] actorIds) {
            // Build Office Graph Query
            string graphQueryText = null;

            if (actorIds.Length > 1) {
                StringBuilder graphQueryBuilder = new StringBuilder();

                foreach (string actorId in actorIds) {
                    if (graphQueryBuilder.Length > 0) {
                        graphQueryBuilder.Append(",");
                    }

                    graphQueryBuilder.AppendFormat("actor({0},action:1020)", actorId);
                }

                graphQueryBuilder.Append(",and(actor(me,action:1021),actor(me,or(action:1021,action:1036,action:1037,action:1039)))");

                graphQueryText = String.Format("or({0})", graphQueryBuilder.ToString());
            }
            else {
                graphQueryText = String.Format("or(actor({0},action:1020),and(actor(me,action:1021),actor(me,or(action:1021,action:1036,action:1037,action:1039))))", actorIds[0]);
            }

            // Ensure that the Web URL is available
            Web web = clientContext.Web;
            if (!web.IsPropertyAvailable("Url")) {
                clientContext.Load(web, w => w.Url);
                clientContext.ExecuteQuery();
            }

            // Configure Search Query
            KeywordQuery searchQuery = new KeywordQuery(clientContext);
            searchQuery.QueryText = String.Format("Path:{0}", web.Url);

            QueryPropertyValue graphQuery = new QueryPropertyValue();
            graphQuery.StrVal = graphQueryText;
            graphQuery.QueryPropertyValueTypeIndex = 1;
            searchQuery.Properties.SetQueryPropertyValue("GraphQuery", graphQuery);

            QueryPropertyValue graphRankingModel = new QueryPropertyValue();
            graphRankingModel.StrVal = @"{""features"":[{""function"":""EdgeWeight""}],""featureCombination"":""sum"",""actorCombination"":""sum""}";
            graphRankingModel.QueryPropertyValueTypeIndex = 1;
            searchQuery.Properties.SetQueryPropertyValue("GraphRankingModel", graphRankingModel);

            searchQuery.SelectProperties.AddRange(new string[] { "Author", "AuthorOwsUser", "DocId", "DocumentPreviewMetadata", "Edges", "EditorOwsUser", "FileExtension", "FileType", "HitHighlightedProperties", "HitHighlightedSummary", "LastModifiedTime", "LikeCountLifetime", "ListID", "ListItemID", "OriginalPath", "Path", "Rank", "SPWebUrl", "SecondaryFileExtension", "ServerRedirectedURL", "SiteTitle", "Title", "ViewCountLifetime", "siteID", "uniqueID", "webID" });
            searchQuery.BypassResultTypes = true;
            searchQuery.RowLimit = 5;
            searchQuery.RankingModelId = "0c77ded8-c3ef-466d-929d-905670ea1d72";
            searchQuery.ClientType = "DocumentsTrendingInThisSite";

            SearchExecutor searchExecutor = new SearchExecutor(clientContext);
            ClientResult<ResultTableCollection> results = searchExecutor.ExecuteQuery(searchQuery);
            clientContext.ExecuteQuery();

            List<TrendingDocument> trendingDocuments = new List<TrendingDocument>(results.Value[0].ResultRows.Count());

            foreach (var row in results.Value[0].ResultRows) {
                string[] lastModifiedByInfo = row["EditorOwsUser"].ToString().Split('|');

                trendingDocuments.Add(new TrendingDocument(
                    row["Title"].ToString(),
                    row["ServerRedirectedURL"].ToString(),
                    GetPreviewImageUrl(row, web.Url),
                    (DateTime)row["LastModifiedTime"],
                    lastModifiedByInfo[1].Trim(),
                    GetUserPhotoUrl(lastModifiedByInfo[0].Trim(), web.Url)));
            }

            return trendingDocuments;
        }

        private static string GetPreviewImageUrl(IDictionary<string, object> resultRow, string webUrl) {
            string uniqueId = resultRow["uniqueID"].ToString();
            string siteId = resultRow["siteID"].ToString();
            string webId = resultRow["webID"].ToString();
            string docId = resultRow["DocId"].ToString();

            return String.Format("{0}/_layouts/15/getpreview.ashx?guidFile={1}&guidSite={2}&guidWeb={3}&docid={4}&metadatatoken=300x424x2&ClientType=DocumentsTrendingInThisSite&size=small",
                webUrl, uniqueId, siteId, webId, docId);
        }

        private static string GetUserPhotoUrl(string userEmail, string webUrl) {
            return String.Format("{0}/_layouts/15/userphoto.aspx?size=S&accountname={1}", webUrl, userEmail);
        }
    }
}

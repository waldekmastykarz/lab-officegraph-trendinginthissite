using Microsoft.SharePoint.Client.Search.Query;

namespace TrendingInThisSiteWeb.Extensions {
    internal static class StringCollectionExtensions {
        internal static void AddRange(this StringCollection stringCollection, string[] arr) {
            if (arr != null) {
                foreach (string s in arr) {
                    stringCollection.Add(s);
                }
            }
        }
    }
}
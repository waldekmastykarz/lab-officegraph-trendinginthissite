using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace TrendingInThisSiteWeb.Models {
    public class TrendingDocument {
        public string Title { get; set; }
        public string Url { get; set; }
        public string PreviewImageUrl { get; set; }
        public DateTime LastModifiedTime { get; set; }
        public string LastModifiedByName { get; set; }
        public string LastModifiedByPhotoUrl { get; set; }

        public TrendingDocument(string title, string url, string previewImageUrl, DateTime lastModifiedTime, string lastModifiedByName, string lastModifiedByPhotoUrl) {
            Title = title;
            Url = url;
            PreviewImageUrl = previewImageUrl;
            LastModifiedTime = lastModifiedTime;
            LastModifiedByName = lastModifiedByName;
            LastModifiedByPhotoUrl = lastModifiedByPhotoUrl;
        }

        public string GetDisplayDate() {
            double differenceInDays = (DateTime.Now - LastModifiedTime).TotalDays;

            string displayDate = null;

            if (differenceInDays > 6) {
                displayDate = LastModifiedTime.ToString("MMMM dd");

                if (differenceInDays > 365) {
                    displayDate = String.Format("{0}, {1}", displayDate, LastModifiedTime.ToString("yyyy"));
                }
            }
            else if (differenceInDays == 1) {
                displayDate = String.Format("Yesterday at {0}", LastModifiedTime.ToString("t"));
            }
            else if (differenceInDays == 0) {
                displayDate = String.Format("Today at {0}", LastModifiedTime.ToString("t"));
            }
            else {
                displayDate = String.Format("{0} at {1}", LastModifiedTime.ToString("dddd"), LastModifiedTime.ToString("t"));
            }

            return displayDate;
        }
    }
}
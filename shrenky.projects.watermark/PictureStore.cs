using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint;
using System.Xml;

namespace shrenky.projects.watermark
{
    internal class PictureStore
    {
        private const string PictureStoreTitle = "ShrenkyProjectsWatermarkPicureStore";

        public void EnsurePictureStore(SPWeb web)
        {
            if (web.Lists.TryGetList(PictureStoreTitle) == null)
            {
                SPSecurity.RunWithElevatedPrivileges(() =>
                {
                    Guid pictureStoreGuid = web.Lists.Add(PictureStoreTitle, "Picture Library to store watermark images", SPListTemplateType.PictureLibrary);
                    web.AllowUnsafeUpdates = true;
                    web.Update();
                    web.AllowUnsafeUpdates = false;
                });
            }
        }

        public bool SavePicture(SPWeb web, string fileName, byte[] content)
        {
            EnsurePictureStore(web);
            bool result = true;
            try
            {
                SPList pictureStore = web.Lists[PictureStoreTitle];
                string destUrl = string.Format("/{0}/{1}", pictureStore.RootFolder.Url, fileName);
                pictureStore.RootFolder.Files.Add(destUrl, content);
                pictureStore.RootFolder.Update();
            }
            catch (Exception)
            {
                result = false;
            }
            return result;
        }

        public ImageProperty GetAllThumbnailLinks(SPWeb web)
        {
            ImageProperty property = new ImageProperty();
            SPList pictureList = web.Lists.TryGetList(PictureStoreTitle);
            if (pictureList != null)
            {
                foreach (SPListItem item in pictureList.Items)
                {
                    property = this.GetImageInfo(item);
                }
            }
            return property;
        }

        #region

        private ImageProperty GetImageInfo(SPListItem item)
        {
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(item.Xml);
            XmlNode row = doc.FirstChild;
            return new ImageProperty()
            {
                Id = item.ID,
                RelativeImageUrl = item.Url,
                ThumbnailExists = row.Attributes["ows_ThumbnailExists"].Value == "1",
                PreviewExists = row.Attributes["ows_PreviewExists"].Value == "1",
                EncodedAbsThumbnailUrl = row.Attributes["ows_EncodedAbsThumbnailUrl"].Value,
                EncodedAbsWebImgUrl = row.Attributes["ows_EncodedAbsWebImgUrl"].Value
            };
        }

        #endregion
    }
}

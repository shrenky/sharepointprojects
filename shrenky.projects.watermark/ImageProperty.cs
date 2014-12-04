using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace shrenky.projects.watermark
{
    internal class ImageProperty
    {
        public int Id { get; set; }
        public string RelativeImageUrl { get; set; }
        //ows_ThumbnailExists
        public bool ThumbnailExists { get; set; }
        //ows_EncodedAbsThumbnailUrl
        public string EncodedAbsThumbnailUrl { get; set; }
        //ows_PreviewExists
        public bool PreviewExists { get; set; }
        //ows_EncodedAbsWebImgUrl
        public string EncodedAbsWebImgUrl { get; set; }
    }
}

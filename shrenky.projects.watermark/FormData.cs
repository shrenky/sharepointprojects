using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace shrenky.projects.watermark
{
    [Serializable]
    public class FormData
    {
        public string WaterMarkText { get; set; }
    }

    public class FormDataHelper
    {
        public static FormData DeserializeFormData(string xmlString)
        {
            using (MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(xmlString)))
            {
                XmlSerializer serializer = new XmlSerializer(typeof(FormData));
                FormData data = (FormData)serializer.Deserialize(stream);
                return data;
            }
        }
    }
}

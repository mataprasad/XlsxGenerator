using ICSharpCode.SharpZipLib.Zip;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;
using System.Xml.Serialization;

namespace XlsxGenerator
{
    public static class Helper
    {
        public const string TEMPLATE_RESOURCE_NAME = "XlsxGenerator.Template.template_without_sheet.xlsx";

        public static Stream GetEmbeddedResourceAsFileStream(string resourceName)
        {
            return Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName);
        }

        public static string SerializeToXML<T>(T obj)
        {
            XmlSerializer serializer = new XmlSerializer(typeof(T));
            XmlSerializerNamespaces ns = new XmlSerializerNamespaces();
            ns.Add(string.Empty, string.Empty);
            using (StringWriter stream = new StringWriter())
            {
                serializer.Serialize(stream, obj);
                stream.Close();
                return stream.ToString();
            }
        }

        public static void CopyStream(Stream input, Stream output)
        {
            byte[] buffer = new byte[8 * 1024];
            int len;
            while ((len = input.Read(buffer, 0, buffer.Length)) > 0)
            {
                output.Write(buffer, 0, len);
            }
        }

        public static void AddFileToExistingZip(string zipFilePath,string pathOfFileToAdd)
        {
            ZipFile zipFile = new ZipFile(zipFilePath);
            zipFile.BeginUpdate();
            zipFile.Add(pathOfFileToAdd, "xl/worksheets/sheet1.xml");
            zipFile.CommitUpdate();
            zipFile.Close();
        }

    }
}

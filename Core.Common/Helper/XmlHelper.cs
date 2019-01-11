using System;
using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace Core.Common.Helper
{
    /// <summary>
    /// XmlHelper
    /// </summary>
    public class XmlHelper
    {
        private static void XmlSerializeInternal(Stream stream, object obj, Encoding encoding, bool isnamespaces)
        {
            if (obj == null)
                throw new ArgumentNullException("o");
            if (encoding == null)
                throw new ArgumentNullException("encoding");

            XmlSerializer serializer = new XmlSerializer(obj.GetType());

            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Indent = true;
            settings.NewLineChars = "\r\n";
            settings.Encoding = encoding;
            settings.IndentChars = "    ";

            //不生成声明头
            settings.OmitXmlDeclaration = !isnamespaces;

            MemoryStream w = new MemoryStream();

            XmlSerializerNamespaces namespaces = new XmlSerializerNamespaces();
            namespaces.Add("", "");

            using (XmlWriter writer = XmlWriter.Create(stream, settings))
            {
                serializer.Serialize(writer, obj, namespaces);
                writer.Close();

            }
        }

        /// <summary>
        /// 将一个对象序列化为XML字符串
        /// </summary>
        /// <param name="obj">要序列化的对象</param>
        /// <param name="encoding">编码方式</param>
        /// <param name="isnamespaces">是否需要命名空间true：需要 false:不需要</param>
        /// <returns>序列化产生的XML字符串</returns>
        public static string EntityToXml(object obj, Encoding encoding, bool isnamespaces)
        {
            using (MemoryStream stream = new MemoryStream())
            {
                XmlSerializeInternal(stream, obj, encoding, isnamespaces);

                stream.Position = 0;
                using (StreamReader reader = new StreamReader(stream, encoding))
                {
                    return reader.ReadToEnd();
                }
            }
        }

        /// <summary>
        /// 从XML字符串中反序列化对象
        /// </summary>
        /// <typeparam name="T">结果对象类型</typeparam>
        /// <param name="xml">包含对象的XML字符串</param>
        /// <param name="encoding">编码方式</param>
        /// <returns>反序列化得到的对象</returns>
        public static T XmlToEntity<T>(string xml, Encoding encoding)
        {
            if (string.IsNullOrEmpty(xml))
                throw new ArgumentNullException("s");
            if (encoding == null)
                throw new ArgumentNullException("encoding");

            XmlSerializer mySerializer = new XmlSerializer(typeof(T));
            using (MemoryStream ms = new MemoryStream(encoding.GetBytes(xml)))
            {
                using (StreamReader sr = new StreamReader(ms, encoding))
                {
                    return (T)mySerializer.Deserialize(sr);
                }
            }
        }

        /// <summary>
        /// 将一个对象按XML序列化的方式写入到一个文件
        /// </summary>
        /// <param name="obj">要序列化的对象</param>
        /// <param name="path">保存文件路径</param>
        /// <param name="encoding">编码方式</param>
        /// <param name="isnamespaces">是否需要命名空间true：需要 false:不需要</param>
        public static void XmlSerializeToFile(object obj, string path, Encoding encoding, bool isnamespaces)
        {
            if (string.IsNullOrEmpty(path))
                throw new ArgumentNullException("path");

            using (FileStream file = new FileStream(path, FileMode.Create, FileAccess.Write))
            {
                XmlSerializeInternal(file, obj, encoding, isnamespaces);
            }
        }

        /// <summary>
        /// 读入一个文件，并按XML的方式反序列化对象。
        /// </summary>
        /// <typeparam name="T">结果对象类型</typeparam>
        /// <param name="path">文件路径</param>
        /// <param name="encoding">编码方式</param>
        /// <returns>反序列化得到的对象</returns>
        public static T XmlDeserializeFromFile<T>(string path, Encoding encoding)
        {
            if (string.IsNullOrEmpty(path))
                throw new ArgumentNullException("path");
            if (encoding == null)
                throw new ArgumentNullException("encoding");

            string xml = File.ReadAllText(path, encoding);
            return XmlToEntity<T>(xml, encoding);
        }
    }
}


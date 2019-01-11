using Newtonsoft.Json;
using System.Collections.Generic;
using System.IO;
using System.Runtime.Serialization.Json;
using System.Text;

namespace Core.Common.Helper
{
    /// <summary>
    /// JsonHelper
    /// </summary>
    public class JsonHelper
    {
        /// <summary>
        /// 类对像转换成json格式
        /// </summary> 
        /// <returns></returns>
        public static string EntityToJson(object obj)
        {
            return JsonConvert.SerializeObject(obj, Newtonsoft.Json.Formatting.Indented, new JsonSerializerSettings { NullValueHandling = NullValueHandling.Include });
        }
        /// <summary>
        /// 类对像转换成json格式
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="HasNullIgnore">是否忽略NULL值</param>
        /// <returns></returns>
        public static string EntityToJson(object obj, bool HasNullIgnore)
        {
            if (HasNullIgnore)
                return JsonConvert.SerializeObject(obj, Newtonsoft.Json.Formatting.Indented, new JsonSerializerSettings { NullValueHandling = NullValueHandling.Ignore });
            else
                return EntityToJson(obj);
        }
        /// <summary>
        /// json格式转换
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="strJson"></param>
        /// <returns></returns>
        public static T JsonToEntity<T>(string strJson) where T : class
        {
            if (!string.IsNullOrEmpty(strJson))
                return JsonConvert.DeserializeObject<T>(strJson);
            return null;
        }
        /// <summary>
        /// json格式转换
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="strJson"></param>
        /// <returns></returns>
        public static List<T> JsonToList<T>(string strJson) where T : class
        {
            if (!string.IsNullOrEmpty(strJson))
                return JsonConvert.DeserializeObject<List<T>>(strJson);
            return null;
        }
        /// <summary>
        /// 功能描述：将List转换为Json
        /// </summary>
        /// <param name="list"></param>
        /// <returns></returns>
        public static string EntityToJson(IList<object> list)
        {
            DataContractJsonSerializer json = new DataContractJsonSerializer(list.GetType());
            string szJson = "";
            //序列化
            using (MemoryStream stream = new MemoryStream())
            {
                json.WriteObject(stream, list);
                szJson = Encoding.UTF8.GetString(stream.ToArray());
            }
            return szJson;
        }
    }
}

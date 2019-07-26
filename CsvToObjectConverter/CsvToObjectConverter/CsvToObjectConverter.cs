
namespace CsvToObjectConverter
{
    #region namespace
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Reflection;
    #endregion

    public static class CsvToObjectConverter
    {
        private readonly static string _oldValueToReplase = @"\,";
        private readonly static string _swrveDefaultValue = @"\N";
        private readonly static string _newValueToReplase = ".";
        public static List<T> DeserializeObjectList<T>(string csvContent) where T : class, new()
        {
            try
            {
                var deserializedObjectList = new List<T>();

                if (string.IsNullOrEmpty(csvContent))
                {
                    return deserializedObjectList;
                }
                string[] separator = new string[] { "\n" };
                string[] fileContent = csvContent.Split(separator, StringSplitOptions.RemoveEmptyEntries);
                var headerContent = fileContent.FirstOrDefault().Split(',');
                Dictionary<string, int> headers = GetHeaderInfo(headerContent);

                var contents = fileContent.Skip(1);
                var propertyInfo = typeof(T).GetRuntimeProperties().ToDictionary(a => a.Name, v => v);
                foreach (var content in contents)
                {
                    var instance = DeserializeObject<T>(headers, propertyInfo, content);
                    if (instance != null)
                        deserializedObjectList.Add(instance);
                }

                return deserializedObjectList;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                throw;
            }
        }

        private static Dictionary<string, int> GetHeaderInfo(string[] headerContent)
        {
            Dictionary<string, int> headers = new Dictionary<string, int>();
            for (int index = 0; index < headerContent.Length; index++)
            {
                headers.Add(headerContent[index], index);
            }

            return headers;
        }

        public static T DeserializeObject<T>(Dictionary<string, int> headers, Dictionary<string, PropertyInfo> instancePropertyDetails, string content) where T : class, new()
        {
            try
            {
                bool isValidObject = false;
                var instance = new T();
                content = content.Replace(_oldValueToReplase, _newValueToReplase);
                var values = content.Split(',');
                foreach (var propertyDetail in instancePropertyDetails)
                {
                    if (headers.ContainsKey(propertyDetail.Key))
                    {
                        if (values[headers[propertyDetail.Key]] == _swrveDefaultValue)
                        {
                            propertyDetail.Value.SetValue(instance, null);
                        }
                        else
                        {
                            string value = string.Empty;
                            if (propertyDetail.Value.PropertyType.IsGenericType && propertyDetail.Value.PropertyType.GetGenericTypeDefinition() == typeof(Nullable<>))
                            {
                                var genericType = propertyDetail.Value.PropertyType.GetGenericArguments()[0];
                                value = values[headers[propertyDetail.Key]];
                                if (genericType == typeof(bool))
                                {
                                    value = value == "1" ? bool.TrueString : bool.FalseString;
                                }
                                propertyDetail.Value.SetValue(instance, Convert.ChangeType(value, genericType), null);
                            }
                            else
                            {
                                value = values[headers[propertyDetail.Key]];
                                if (propertyDetail.Value.PropertyType == typeof(bool))
                                {
                                    value = value == "1" ? bool.TrueString : bool.FalseString;
                                }
                                propertyDetail.Value.SetValue(instance, Convert.ChangeType(value, propertyDetail.Value.PropertyType), null);
                            }
                        }
                        isValidObject = true;
                    }
                }
                if (isValidObject)
                    return instance;
                else
                    return null;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                throw;
            }
        }
    }
}

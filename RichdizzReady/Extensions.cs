using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace RichdizzReady
{
	public static class Extensions
	{
		public static async Task<JArray> GetJArray(this HttpClient client, string endpoint, string token)
		{
			using (var response = await client.GetAsync(endpoint))
			{
				if (response.IsSuccessStatusCode)
				{
					var json = await response.Content.ReadAsStringAsync();
					return (JArray)JObject.Parse(json).SelectToken("value");
				}
				else
				    return null;
			}
		}

		public static async Task<string> GetJArrayPaged(this HttpClient client, string endpoint, JArray array)
		{
			using (var response = await client.GetAsync(endpoint))
			{
                if (response.IsSuccessStatusCode)
                {
                    var text = await response.Content.ReadAsStringAsync();
                    var json = JObject.Parse(text);
                    JArray newItems = (JArray)json.SelectToken("value");
                    array.AddRange(newItems);

                    if (json["@odata.nextLink"] != null)
                    {
                        // Recursively get more
                        return await client.GetJArrayPaged(json["@odata.nextLink"].Value<string>(), array);
                    }
                    else if (json["@odata.deltaLink"] != null)
                    {
                        return json["@odata.deltaLink"].Value<string>();
                    }
                    else
                        return null; // ERROR
                }
                else
                    return null; // ERROR
			}
		}

		public static void AddRange(this JArray array, JArray jarray)
		{
			foreach (var token in jarray)
				array.Add(token);
		}
	}
}

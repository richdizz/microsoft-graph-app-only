using System;
using System.Threading.Tasks;
using MongoDB.Driver;
using MongoDB.Bson.Serialization.Attributes;
using MongoDB.Bson;
using System.Security.Authentication;
using System.Net.Http;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System.Security.Cryptography.X509Certificates;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;
using Newtonsoft.Json;
using System.Text;

namespace RichdizzReady
{
    class Program
    {
        static void Main(string[] args)
        {
			var task = RunProcess();
			Task.WaitAll(task);
        }

        private static async Task RunProcess()
        {
			// Establish connection to CosmosDB
			string connectionString = @"MONGO-CONNECTION-STRING-HERE";
			MongoClientSettings settings = MongoClientSettings.FromUrl(new MongoUrl(connectionString));
			settings.SslSettings = new SslSettings() { EnabledSslProtocols = SslProtocols.Tls12 };
			var mongoClient = new MongoClient(settings);
			var db = mongoClient.GetDatabase("richdizzready");
			var collection = db.GetCollection<User>("users");

			// Query all users that we have github tokens for
			var users = await collection.Find(i => i.alias != null).ToListAsync();
            foreach (var user in users)
            {
                Console.WriteLine($"Processing {user.alias}");

				// Read the certificate private key from the executing location
				// NOTE: This is a hack...Azure Key Vault is best approach...also certPath format will vary by platform
				var certPath = System.Reflection.Assembly.GetEntryAssembly().Location;
				certPath = certPath.Substring(0, certPath.LastIndexOf("/bin", StringComparison.CurrentCultureIgnoreCase)) + "/RichdizzReadyKey.pfx";
				var certfile = System.IO.File.OpenRead(certPath);
				var certificateBytes = new byte[certfile.Length];
				certfile.Read(certificateBytes, 0, (int)certfile.Length);
				var cert = new X509Certificate2(
					certificateBytes,
					"P@ssword",
					X509KeyStorageFlags.Exportable |
					X509KeyStorageFlags.MachineKeySet |
					X509KeyStorageFlags.PersistKeySet); //switchest are important to work in webjob
				ClientAssertionCertificate cac = new ClientAssertionCertificate("f0dc44ea-333f-4060-9c89-a7730f9b92fe", cert);

				// Get the access token to the Microsoft Graph using the ClientAssertionCertificate
				Console.WriteLine("Getting app-only access token to Microsoft Graph");
				string authority = "https://login.microsoftonline.com/rzna.onmicrosoft.com/"; // Currently hard-coded to be single-tenant
				AuthenticationContext authenticationContext = new AuthenticationContext(authority, false);
				var authenticationResult = await authenticationContext.AcquireTokenAsync("https://graph.microsoft.com", cac);
				var token = authenticationResult.AccessToken;
				Console.WriteLine("App-only access token retreived");

                // Check for a delta token which would indicate if we have ever processed this user
                if (String.IsNullOrEmpty(user.delta_token))
                {
					// This is a new user so we need to get their sent items folder id and process items
                    string endpoint = $"https://graph.microsoft.com/v1.0/users/{user.alias}/mailfolders";
					HttpClient foldersClient = new HttpClient();
					foldersClient.DefaultRequestHeaders.Add("Accept", "application/json");
					foldersClient.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);
                    var folders = await foldersClient.GetJArray(endpoint, token);

                    // Get the "Sent Items" folder for the user
                    string sentItemsFolderId = String.Empty;
                    foreach (var folder in folders)
                    {
                        if (folder["displayName"].Value<string>() == "Sent Items")
                        {
                            sentItemsFolderId = folder["id"].Value<string>();
                            break;
                        }
                    }
                    user.delta_token = $"https://graph.microsoft.com/v1.0/users/{user.alias}/mailfolders('{sentItemsFolderId}')/messages/delta?$select=id,body";
                }

                // Prepare the request to recursively get messages in the sent folder
				HttpClient client = new HttpClient();
				client.DefaultRequestHeaders.Add("Accept", "application/json");
				client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);
				client.DefaultRequestHeaders.Add("Prefer", "outlook.body-content-type=\"text\""); // We want body as text
				JArray messages = new JArray();
				var deltaToken = await client.GetJArrayPaged(user.delta_token, messages);

                // Build up a payload of all messages so we can make just one call into the text analytics service
                var payload = new Documents();
                for (var i = 0; i < messages.Count; i++)
                {
                    payload.documents.Add(new Document() { id = i, language = "en", text = messages[i].SelectToken("body.content").Value<string>() });
                }

                // Only try to get sentiment if we are processing one or more messages
                if (payload.documents.Count > 0)
                {
                    // Get sentiment for all the messages
                    HttpClient sentimentClient = new HttpClient();
                    sentimentClient.DefaultRequestHeaders.Add("Accept", "application/json");
                    sentimentClient.DefaultRequestHeaders.Add("Ocp-Apim-Subscription-Key", "03c751a65dcd4263a8cec60be5fdcfa5");
                    StringContent content = new StringContent(JsonConvert.SerializeObject(payload), Encoding.UTF8, "application/json");
                    using (var resp = await sentimentClient.PostAsync("https://westus.api.cognitive.microsoft.com/text/analytics/v2.0/sentiment", content))
                    {
                        if (resp.IsSuccessStatusCode)
                        {
                            // Add the sentiment as an open extension on the message
                            var sentimentResponse = await resp.Content.ReadAsStringAsync();
                            var docs = (JArray)JObject.Parse(sentimentResponse)["documents"];
                            for (var j = 0; j < messages.Count; j++)
                            {
                                string openExtension = @"
	                            {
								    '@odata.type':'microsoft.graph.openTypeExtension',
								    'extensionName':'com.richdizz.sentiment',
								    'sentiment':'" + docs[j].SelectToken("score").Value<Decimal>() + @"'
								}";

                                client = new HttpClient();
                                client.DefaultRequestHeaders.Add("Accept", "application/json");
                                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);
                                StringContent extContent = new StringContent(openExtension, Encoding.UTF8, "application/json");
                                var extEndpoint = $"https://graph.microsoft.com/v1.0/users/{user.alias}/messages/{messages[j].SelectToken("id").Value<String>()}/extensions";
                                using (var extResp = await client.PostAsync(extEndpoint, extContent))
                                {
                                    // throw exception if open extension failed
                                    if (!extResp.IsSuccessStatusCode)
                                        throw new Exception(await extResp.Content.ReadAsStringAsync());
                                }
                            }

                            // Save the ner delta token for the user
                            user.last_run = DateTime.UtcNow;
                            var filter = Builders<User>.Filter.Eq(u => u.alias, user.alias);
                            await collection.ReplaceOneAsync(filter, user);
                        }
                        else
                        {
                            // Do not store the new deltalink...let it reprocess
                        }
                    }
                }
                else
                {
                    // Save the ner delta token for the user
                    user.last_run = DateTime.UtcNow;
					var filter = Builders<User>.Filter.Eq(u => u.alias, user.alias);
					await collection.ReplaceOneAsync(filter, user);
                }

                Console.WriteLine($"Complete {user.alias}");
            }

            Console.WriteLine($"Processing complete");
        }
    }

	public class User
	{
		[BsonId]
		public ObjectId _id { get; set; }
		public string alias { get; set; }
        public string delta_token { get; set; }
        public DateTime last_run { get; set; }
	}

    public class Documents
    {
        public Documents() { documents = new List<Document>(); }
        public List<Document> documents { get; set; }
    }
    public class Document
    {
        public string language { get; set; }
		public int id { get; set; }
		public string text { get; set; }
    }
}

using Microsoft.PowerPlatform.Dataverse.Client;
using Microsoft.Xrm.Tooling.Connector;
using Microsoft.Xrm.Sdk;

class Program
{
    static void Main(string[] args)
    {
        // Connection string for connecting to Dataverse
        string connectionString = "AuthType=ClientSecret;Url=<your_dataverse_url>;ClientId=<your_client_id>;ClientSecret=<your_client_secret>";

        // Create a CrmServiceClient object using the connection string
        CrmServiceClient service = new CrmServiceClient(connectionString);

        // Check if connection is successful
        if (service != null && service.IsReady)
        {
            // Create an Entity object for the record you want to create
            Entity newRecord = new Entity("your_entity_logical_name");
            
            // Set attributes for the new record
            newRecord["attribute1_logical_name"] = "value1";
            newRecord["attribute2_logical_name"] = 123;
            // Add more attributes as needed

            // Create the record in Dataverse
            Guid newRecordId = service.Create(newRecord);

            Console.WriteLine("Record created successfully with ID: " + newRecordId);
        }
        else
        {
            Console.WriteLine("Failed to connect to Dataverse.");
        }
    }
}

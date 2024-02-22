using Microsoft.PowerPlatform.Dataverse.Client;
using Microsoft.Xrm.Sdk;
using System;

public class Program
{
    static void Main(string[] args)
    {
        // Connection string for connecting to Dataverse
        string connectionString = "AuthType=ClientSecret;Url=<your_dataverse_url>;ClientId=<your_client_id>;ClientSecret=<your_client_secret>";

        // Create an instance of DataverseOperations
        DataverseOperations dataverseOperations = new DataverseOperations(connectionString);

        // Example usage: Create a record
        Entity newRecord = new Entity("your_entity_logical_name");
        newRecord["attribute1_logical_name"] = "value1";
        newRecord["attribute2_logical_name"] = 123;
        Guid newRecordId = dataverseOperations.CreateRecord("your_entity_logical_name", newRecord);

        // Example usage: Retrieve a record
        if (newRecordId != Guid.Empty)
        {
            Entity retrievedRecord = dataverseOperations.RetrieveRecord("your_entity_logical_name", newRecordId, new ColumnSet(true));
        }

        // Example usage: Update a record
        if (newRecordId != Guid.Empty)
        {
            newRecord["attribute1_logical_name"] = "updated_value";
            dataverseOperations.UpdateRecord(newRecord);
        }

        // Example usage: Delete a record
        if (newRecordId != Guid.Empty)
        {
            dataverseOperations.DeleteRecord("your_entity_logical_name", newRecordId);
        }
    }

    public class DataverseOperations
    {
        private CrmServiceClient _service;

        // Constructor to initialize the CrmServiceClient
        public DataverseOperations(string connectionString)
        {
            _service = new CrmServiceClient(connectionString);
        }

        // Create a record in Dataverse
        public Guid CreateRecord(string entityLogicalName, Entity record)
        {
            if (_service != null && _service.IsReady)
            {
                Guid newRecordId = _service.Create(record);
                Console.WriteLine("Record created successfully with ID: " + newRecordId);
                return newRecordId;
            }
            else
            {
                Console.WriteLine("Failed to connect to Dataverse.");
                return Guid.Empty;
            }
        }

        // Retrieve a record from Dataverse
        public Entity RetrieveRecord(string entityLogicalName, Guid recordId, ColumnSet columns)
        {
            if (_service != null && _service.IsReady)
            {
                Entity retrievedRecord = _service.Retrieve(entityLogicalName, recordId, columns);
                Console.WriteLine("Record retrieved successfully.");
                return retrievedRecord;
            }
            else
            {
                Console.WriteLine("Failed to connect to Dataverse.");
                return null;
            }
        }

        // Update a record in Dataverse
        public void UpdateRecord(Entity record)
        {
            if (_service != null && _service.IsReady)
            {
                _service.Update(record);
                Console.WriteLine("Record updated successfully.");
            }
            else
            {
                Console.WriteLine("Failed to connect to Dataverse.");
            }
        }

        // Delete a record from Dataverse
        public void DeleteRecord(string entityLogicalName, Guid recordId)
        {
            if (_service != null && _service.IsReady)
            {
                _service.Delete(entityLogicalName, recordId);
                Console.WriteLine("Record deleted successfully.");
            }
            else
            {
                Console.WriteLine("Failed to connect to Dataverse.");
            }
        }
    }
}

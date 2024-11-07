// Import modules
import "isomorphic-fetch";
import { ClientSecretCredential } from '@azure/identity';
import { Client } from '@microsoft/microsoft-graph-client';
import { TokenCredentialAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js';

// Define variables
let _settings = undefined;
let _clientSecret = undefined;
let _appClient = undefined;

// Initialize Graph authentication (app only)
export async function initializeGraph(settings) {
    // Check that settings are defined and if not - define them
    if (!settings){throw new Error("[ERROR] Settings are undefined.");}
    _settings = settings;

    // Define client secret
    if (!_clientSecret) {
        _clientSecret = new ClientSecretCredential(
            _settings.tenantId,
            _settings.clientId,
            _settings.clientSecret,
        );
    }

    // Define app client
    if (!_appClient) {
        const authProvider = new TokenCredentialAuthenticationProvider(
            _clientSecret,
            {
                scopes: ["https://graph.microsoft.com/.default"],
            }
        );

        _appClient = Client.initWithMiddleware({
            authProvider: authProvider,
        });
    }
}

export async function fetchDeviceAndUsers(serials) {
    const device = await fetchDevices(serials);

    console.log("Fetched data:", device[0].deviceName);
    console.log("Fetched data:", device[1].deviceName);

    let result = [];

    
    for (let i = 0; i < (device.length); i++) {
        result.push({
            deviceName: device[i].deviceName,
        });
    }

    if (device) {
        JSON.stringify(result);
        console.log("After JSON:", result);
    }

    return result;
}

async function fetchDevices(serial) {

    var serial_array = serial.split("\n");
    var result = [];
    console.log("Searching device with serial", serial);

    for (const obj in serial_array) {
        const device = await _appClient
            .api("/deviceManagement/managedDevices")
            .version("beta")
            .filter(`contains(deviceName, '${serial_array[obj]}')`)
            .select(["deviceName", "id", "usersLoggedOn"])
            .get();

        console.log("Inside function:", device.value[0].deviceName);
        result.push({
            deviceName: device.value[0].deviceName,
        });
    }

    return result;

    /*
    const result = await _appClient
        .api("/deviceManagement/managedDevices")
        .version("beta")
        .filter(`contains(deviceName, '${serial}')`)
        .select(["deviceName", "id", "usersLoggedOn"])
        .get();

    console.log("Found", result.value[0]);
    console.log("deviceName:", result.value[0].deviceName);
    console.log("id:", result.value[0].id);

    return result.value;
    */
}
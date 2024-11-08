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
    const device = await fetchData(serials);

    let result = [];

    for (let i = 0; i < (device.length); i++) {
        //console.log("Fetched data:", device[i].deviceName + device[i].azureADDeviceId);
        result.push({
            deviceName: device[i].deviceName,
            azureADDeviceId: device[i].azureADDeviceId
        });
    }

    if (device) {
        JSON.stringify(result);
        console.log("After JSON:", result);
    }

    return result;
}

// Fetch device and check if they exist. Save identifying data.
async function fetchData(serial) {

    var serial_array = serial.split("\n");
    var result = [];
    console.log("Searching device with serial", serial);

    var intune = [];
    var autopilot = [];
    var entra = [];

    //intune = fetchIntune(serial_array);
    autopilot = fetchAutopilot(serial_array);
    entra = fetchEntra(serial_array);

    result = fetchIntune(serial_array);
    //fetchAutopilot(serial_array);
    //fetchEntra(serial_array)

    return result;
}

// Intune check
async function fetchIntune(serial_array) {

    var result = [];

    for (const obj in serial_array) {
        const device = await _appClient
            .api("/deviceManagement/managedDevices")
            .version("beta")
            .filter(`contains(serialNumber, '${serial_array[obj]}')`)
            // Optional optimising (for later)
            //.select(["deviceName", "id", "usersLoggedOn", "azureADDeviceId"])
            .get();

        if (device.value[0] === undefined){
            console.log("No device for", serial_array[obj] ,"found in Intune.");
        } else {
            console.log("Device", serial_array[obj] ,"found from Intune");
            result.push({
                deviceName: device.value[0].deviceName,
                azureADDeviceId: device.value[0].azureADDeviceId
            });
        }
    }

    return result;
}

// Autopilot check
async function fetchAutopilot(serial_array) {

    var result = [];

    for (const obj in serial_array) {
        const device = await _appClient
            .api("/deviceManagement/windowsAutopilotDeviceIdentities")
            .filter(`contains(serialNumber, '${serial_array[obj]}')`)
            // Optional optimising (for later)
            //.select(["displayName"])
            .get();

        if (device.value[0] === undefined){
            console.log("No device for", serial_array[obj] ,"found in Autopilot.");
        } else {
            console.log("Device", serial_array[obj] ,"found from Autopilot");
            //result.push({
                //deviceName: device.value[0].deviceName,
                //azureADDeviceId: device.value[0].azureADDeviceId
            //});
        }
    }

    return result;
}

// Entra check
async function fetchEntra(serial_array) {

    var result = [];

    for (const obj in serial_array) {
        const device = await _appClient
            .api("/devices")
            .header("ConsistencyLevel", "eventual")
            .search(`"displayName:${serial_array[0]}"`)
            .get();

        if (device.value[0] === undefined){
            console.log("No device for", serial_array[obj] ,"found in Entra.");
        } else {
            console.log("Device", serial_array[obj] ,"found from Entra");
            //result.push({
                //deviceName: device.value[0].deviceName,
                //azureADDeviceId: device.value[0].azureADDeviceId
            //});
        }
    }

    return result;
}
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
            azureADDeviceId: device[i].id
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
    var temp = [];
    console.log("Searching device with serial", serial);

    var intune = [];
    var autopilot = [];
    var entra = [];

    //intune = fetchIntune(serial_array);
    //result = await fetchIntune(serial_array);

    intune = await fetchIntune(serial_array);
    autopilot = await fetchAutopilot(serial_array);
    entra = await fetchEntra(serial_array);
    
    var i = 0;
    temp = {
        lastLogin: intune[i].lastLogOnDateTime,	                // Latest login
        lastLogOnUser: intune[i].lastLogOnUser,                 // Latest user
        lastLogOnUserEmail: intune[i].lastLogOnUserEmail,       // Latest email
        userPrincipalName: intune[i].userPrincipalName,     	// Entra ID addr
        intuneResult: intune[i].resultId,						// Identifier ID Intune
        intuneFound: intune[i].foundIntune,						// Found Intune?
        intuneName: intune[i].deviceName,						// Device name Intune
        intuneId: intune[i].id,									// Device ID Intune
        autopilotResult: autopilot[i].resultId,					// Identifier ID Autopilot
        autopilotFound: autopilot[i].foundAutopilot,			// Found Autopilot?
        autopilotName: autopilot[i].displayName,				// Display name Autopilot
        autopilotId: autopilot[i].id,							// Device ID Autopilot
        entraResult: entra[i].resultId,							// Identifier ID Entra
        entraFound: entra[i].foundEntra,						// Found Entra?
        entraName: entra[i].displayName,					    // Device name Entra
        entraId: entra[i].deviceId		                        // Device ID Entra
    }						

    console.log("Compiled data", temp);

    //return result;
}

// Intune data
async function fetchIntune(serial_array) {

    var result = [];
    var user = [];
    var foundIntune, deviceName, id, userPrincipalName, usersLoggedOn, lastLogOnUser, lastLogOnUserEmail;

    for (const obj in serial_array) {
        const device = await _appClient
            .api("/deviceManagement/managedDevices")
            .version("beta")
            .filter(`contains(serialNumber, '${serial_array[obj]}')`)
            .select(["deviceName", "id", "usersLoggedOn", "userPrincipalName"])
            .get();

        foundIntune = true;
        if (device.value[0] === undefined){
            console.log("No device for", serial_array[obj] ,"found in Intune.");
            foundIntune = false;
        } else { console.log(serial_array[obj], "found.");
        }

        deviceName = device.value[0].deviceName || "No data";
        id = device.value[0].id || "No data";
        userPrincipalName = device.value[0].userPrincipalName || "No data";
        usersLoggedOn = device.value[0].usersLoggedOn[(device.value[0].usersLoggedOn.length - 1)] || "No data";

        user = await fetchUser(usersLoggedOn.userId);

        result.push({
            resultId: [obj],
            foundIntune: foundIntune,
            deviceName: deviceName,
            id: id,
            userPrincipalName: userPrincipalName,
            lastLogOnDateTime: usersLoggedOn.lastLogOnDateTime,
            lastLogOnUser: user.displayName,
            lastLogOnUserEmail: user.userPrincipalName
        })
    }

    return result;
}

// Autopilot data
async function fetchAutopilot(serial_array) {

    var result = [];
    var foundAutopilot, displayName, id;

    for (const obj in serial_array) {
        const device = await _appClient
            .api("/deviceManagement/windowsAutopilotDeviceIdentities")
            .filter(`contains(serialNumber, '${serial_array[obj]}')`)
            // Optional optimising (for later)
            //.select(["displayName"])
            .get();

        foundAutopilot = true;
        if (device.value[0] === undefined){
            console.log("No device for", serial_array[obj] ,"found.");
            foundAutopilot = false;
        } else { console.log(serial_array[obj], "found.");}

        displayName = device.value[0].displayName || "No data";
        id = device.value[0].id || "No data";

        result.push({
            resultId: [obj],
            foundAutopilot: foundAutopilot,
            displayName: displayName,
            id: id,
        })
    }

    return result;
}

// Entra data
async function fetchEntra(serial_array) {

    var result = [];
    var foundEntra, displayName, deviceId;

    for (const obj in serial_array) {
        const device = await _appClient
            .api("/devices")
            .header("ConsistencyLevel", "eventual")
            .search(`"displayName:${serial_array[0]}"`)
            .get();

        foundEntra = true;
        if (device.value[0] === undefined){
            console.log("No device for", serial_array[obj] ,"found.");
            foundEntra = false;
        } else { console.log(serial_array[obj], "found.");}

        displayName = device.value[0].displayName || "No data";
        deviceId = device.value[0].id || "No data";

        result.push({
            resultId: [obj],
            foundEntra: foundEntra,
            displayName: displayName,
            deviceId: deviceId,
        });
    }

    return result;
}

async function fetchUser(user_id){
    var user = "No data";
    try {
        const user = await _appClient
            .api(`/users/${user_id}`)
            .select(["displayName", "userPrincipalName"])
            .get();

        return user;
    }
    catch (error) {
        console.log("[ERROR] Error fetching user for ID:", user_id, error);
        return user;
    }
}
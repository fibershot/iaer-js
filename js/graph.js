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

// Fetch device and check if it exists. Save identifying data.
export async function fetchData(serial) {

    // The exported serials from app.js - split from the point of linebreak (\n)
    var serial_array = serial.split("\n");
    var intune = [], autopilot = [], entra = [], result = [];

    // Each platform has it's own function for data fetching
    intune = await fetchIntune(serial_array);
    autopilot = await fetchAutopilot(serial_array);
    entra = await fetchEntra(serial_array);
    
    // Merge data from all the arrays into a final version
    for (let i = 0; i < (serial_array.length); i++) {
        result.push({
            lastLogin: intune[i].lastLogOnDateTime, // Latest login
            lastLogOnUser: intune[i].lastLogOnUser, // Latest user
            lastLogOnUserEmail: intune[i].lastLogOnUserEmail,   // Latest email
            userPrincipalName: intune[i].userPrincipalName, // Entra ID addr
            intuneResult: i,   // Identifier ID Intune
            intuneFound: intune[i].foundIntune, // Found Intune?
            intuneName: intune[i].deviceName,   // Device name Intune
            intuneId: intune[i].id, // Device ID Intune
            autopilotResult: i, // Identifier ID Autopilot
            autopilotFound: autopilot[i].foundAutopilot,    // Found Autopilot?
            autopilotName: autopilot[i].displayName,    // Display name Autopilot
            autopilotId: autopilot[i].id,   // Device ID Autopilot
            entraResult: i, // Identifier ID Entra
            entraFound: entra[i].foundEntra,    // Found Entra?
            entraName: entra[i].displayName,    // Device name Entra
            entraId: entra[i].deviceId,  // Device ID Entra
            model: intune[i].model,     // Model number
            lastContactedDateTime: autopilot[i].lastContactedDateTime, // Last contact date time
            // Debug data
            intuneSerial: intune[i].serialNumber,
            autopilotSerial: autopilot[i].serialNumber
        });
    }

    if (result){
        JSON.stringify(result);
        console.log("[Info] Data coverted to JSON.");
        console.log(result);
    }

    return result;
}

// Intune data
async function fetchIntune(serial_array) {

    var result = [];
    var user = [];
    var deviceName, id, userPrincipalName, usersLoggedOn, model, serialNumber, lastLogOnDateTime, lastLogOnUser, lastLogOnUserEmail;

    // api fetch call
    for (const obj in serial_array) {
        const device = await _appClient
            .api("/deviceManagement/managedDevices")
            .version("beta")
            .filter(`contains(serialNumber, '${serial_array[obj]}')`)
            .select(["deviceName", "id", "usersLoggedOn", "userPrincipalName", "model", "serialNumber"])
            .get();

        if (device.value[0] === undefined){
            console.log("[Intune] No device for", serial_array[obj] ,"found in Intune.");
            result.push({
                foundIntune: false,
                deviceName: "No data",
                id: "No data",
                userPrincipalName: "No data",
                lastLogOnDateTime: "No data",
                lastLogOnUser: "No data",
                lastLogOnUserEmail: "No data",
                model: "No data",
                serialNumber: "No data"
            });

            return result;
            
        } else { 
        console.log("[Intune]", serial_array[obj], "found.");

        // Either fill variables with data and if there's none, resort to "No data"
        deviceName = device.value[0].deviceName;
        id = device.value[0].id;
        userPrincipalName = device.value[0].userPrincipalName || "No data";
        serialNumber = device.value[0].serialNumber;
        model = device.value[0].model;

        if (device.value[0].usersLoggedOn.length <= 0){
            lastLogOnDateTime = "No data";
            lastLogOnUser = "No data";
            lastLogOnUserEmail = "No data";
        } else {
            usersLoggedOn = device.value[0].usersLoggedOn[(device.value[0].usersLoggedOn.length - 1)]
            user = await fetchUser(usersLoggedOn.userId);
            lastLogOnDateTime = usersLoggedOn.lastLogOnDateTime;
            lastLogOnUser = user.displayName;
            lastLogOnUserEmail = user.userPrincipalName;
        }


        // Add results to array and return
        result.push({
            foundIntune: true,
            deviceName: deviceName,
            id: id,
            userPrincipalName: userPrincipalName,
            lastLogOnDateTime: lastLogOnDateTime,
            lastLogOnUser: lastLogOnUser,
            lastLogOnUserEmail: lastLogOnUserEmail,
            model: model,
            serialNumber: serialNumber
        });
    }

    return result;
}}

// Autopilot data
async function fetchAutopilot(serial_array) {

    var result = [];
    var displayName, id, lastContactedDateTime, serialNumber;

    // api fetch call
    for (const obj in serial_array) {
        const device = await _appClient
            .api("/deviceManagement/windowsAutopilotDeviceIdentities")
            .filter(`contains(serialNumber, '${serial_array[obj]}')`)
            // Optional optimising (for later)
            //.select(["displayName"])
            .get();

        if (device.value[0] === undefined){
            console.log("[Autopilot] No device for", serial_array[obj] ,"found.");
           
            result.push({
                foundAutopilot: false,
                displayName: "No data",
                id: "No data",
                lastContactedDateTime: "No data",
                serialNumber: "No data"
            });

            return result;
        } 
        else { 
            
        console.log("[Autopilot]", serial_array[obj], "found.");

        // Either fill variables with data and if there's none, resort to "No data"
        displayName = device.value[0].displayName;
        id = device.value[0].id;
        lastContactedDateTime = device.value[0].lastContactedDateTime;
        serialNumber = device.value[0].serialNumber;

        // Add results to array and return
        result.push({
            foundAutopilot: true,
            displayName: displayName,
            id: id,
            lastContactedDateTime: lastContactedDateTime,
            serialNumber: serialNumber
        });
    }

    return result;
}}

// Entra data
async function fetchEntra(serial_array) {

    var result = [];
    var displayName, deviceId;

    // api fetch call
    for (const obj in serial_array) {
        const device = await _appClient
            .api("/devices")
            .header("ConsistencyLevel", "eventual")
            .search(`"displayName:${serial_array[obj]}"`)
            .get();

        if (device.value[0] === undefined){
            console.log("[Entra] No device for", serial_array[obj] ,"found.");
            result.push({
                foundEntra: false,
                displayName: "No data",
                deviceId: "No data"
            });

            return result;

        } else { 
            
        console.log("[Entra]", serial_array[obj], "found.");

        // Either fill variables with data and if there's none, resort to "No data"
        displayName = device.value[0].displayName;
        deviceId = device.value[0].id;

        // Add results to array and return
        result.push({
            foundEntra: true,
            displayName: displayName,
            deviceId: deviceId,
        });

        return result;
    }
}}

// Fetch user from the lastest user ID that logged in
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
        console.log("[ERROR] Error fetching user for ID:", user_id);
        return user;
    }
}

// Function for deletion
export async function removeDevice(serial) {
    var serial_array = serial.split("\n");
    console.log(serial_array[0], serial_array[1], serial_array[2]);
}

async function deleteIntune(DEVICE_ID) {

}

async function deleteAutopilot(DEVICE_ID) {
    
}

async function deleteEntra(DEVICE_ID) {

}
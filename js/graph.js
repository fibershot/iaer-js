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
            intuneSerial: intune[i].serialNumber, // Intune serial
            model: intune[i].model,     // Model number

            autopilotResult: i, // Identifier ID Autopilot
            autopilotFound: autopilot[i].foundAutopilot,    // Found Autopilot?
            autopilotName: autopilot[i].displayName,    // Display name Autopilot
            autopilotId: autopilot[i].id,   // Device ID Autopilot
            autopilotSerial: autopilot[i].serialNumber, // Autopilot serial
            lastContactedDateTime: autopilot[i].lastContactedDateTime, // Last contact date time

            entraResult: i, // Identifier ID Entra
            entraFound: entra[i].foundEntra,    // Found Entra?
            entraName: entra[i].displayName,    // Device name Entra
            entraId: entra[i].deviceId,  // Device ID Entra

            initialReq: serial_array[i],
        });
    }

    if (result){
        JSON.stringify(result);
        console.log("[!] Data coverted to JSON", result.length, "results.");
    }

    return result;
}

// Intune data
async function fetchIntune(serial_array) {
    var result = [];
    var data, deviceName, id, userPrincipalName, usersLoggedOn, model, serialNumber, lastLogOnDateTime, lastLogOnUser, lastLogOnUserEmail;

    // API fetch request
    for (const obj in serial_array) {
        const device = await _appClient
        .api("/deviceManagement/managedDevices")
        .version("beta")
        .filter(`contains(serialNumber, '${serial_array[obj]}')`)
        .select(["deviceName", "id", "usersLoggedOn", "userPrincipalName", "model", "serialNumber"])
        .get();

        data = device.value[0];

        // Assign data to variables (normal data)
        deviceName = data?.deviceName || "Unknown";
        id = data?.id || "Unknown";
        userPrincipalName = data?.userPrincipalName || "Unknown";
        model = data?.model || "Unknown";
        serialNumber = data?.serialNumber || "Unknown";
        usersLoggedOn = data?.usersLoggedOn || "Unknown"; // Not in return
        lastLogOnUser = "Unknown";
        lastLogOnUserEmail = "Unknown";
        lastLogOnDateTime = "Unknown";

        // Check if any users have logged in recently
        if (usersLoggedOn.length > 0) {
            usersLoggedOn = data?.usersLoggedOn[(data.usersLoggedOn.length - 1)] || "Unknown";
            lastLogOnDateTime = usersLoggedOn?.lastLogOnDateTime || "Unknown";
            const user = await fetchUser(usersLoggedOn.userId);
            lastLogOnUser = user?.displayName || "Unknown";
            lastLogOnUserEmail = user?.userPrincipalName || "Unknown";
        }

        // Push results
        result.push({
            foundIntune: device?.value[0] ? true : false,
            deviceName: deviceName,
            id: id,
            userPrincipalName: userPrincipalName,
            model: model,
            serialNumber: serialNumber,
            lastLogOnDateTime: lastLogOnDateTime,
            lastLogOnUser: lastLogOnUser,
            lastLogOnUserEmail: lastLogOnUserEmail
        });

        if (result[obj].foundIntune) {
            console.log("[I] Fetched", serial_array[obj], "from Intune.");
        } else {
            console.log("[I] Failed fetch for", serial_array[obj], "from Intune.");
        }
    }

    return result;
}

// Autopilot data
async function fetchAutopilot(serial_array) {
    var result = [];
    var data, displayName, id, lastContactedDateTime, serialNumber;

    // API fetch call
    for (const obj in serial_array) {
        const device = await _appClient
        .api("/deviceManagement/windowsAutopilotDeviceIdentities")
        .filter(`contains(serialNumber, '${serial_array[obj]}')`)
        .get();

        data = device.value[0];

        // Assign data to variables
        displayName = data?.displayName || "Unknown";
        id = data?.id || "Unknown";
        lastContactedDateTime = data?.lastContactedDateTime || "Unknown";
        serialNumber = data?.serialNumber || "Unknown";

        // Add results to array and return
        result.push({
            foundAutopilot: device?.value[0] ? true : false,
            displayName: displayName,
            id: id,
            lastContactedDateTime: lastContactedDateTime,
            serialNumber: serialNumber
        });

        if (result[obj].foundAutopilot) {
            console.log("[A] Fetched", serial_array[obj], "from Autopilot.");
        } else {
            console.log("[A] Failed fetch for", serial_array[obj], "from Autopilot.");
        }
    }
    return result;
}

// Entra data
async function fetchEntra(serial_array) {
    var result = [];
    var data, displayName, deviceId;

    // API fetch call
    for (const obj in serial_array) {
        const device = await _appClient
        .api("/devices")
        .header("ConsistencyLevel", "eventual")
        .search(`"displayName:${serial_array[obj]}"`)
        .get();

        data = device.value[0];

        displayName = data?.displayName || "Unknown";
        deviceId = data?.id || "Unknown";

        result.push({
            foundEntra: device?.value[0] ? true : false,
            displayName: displayName,
            deviceId: deviceId
        });

        if (result[obj].foundEntra) {
            console.log("[E] Fetched", serial_array[obj], "from Entra.");
        } else {
            console.log("[E] Failed fetch for", serial_array[obj], "from Entra.");
        }
    }
    return result;
}

// Fetch user from the lastest user ID that logged in
async function fetchUser(user_id){
    try {
        const user = await _appClient
            .api(`/users/${user_id}`)
            .select(["displayName", "userPrincipalName"])
            .get();

        console.log("[U] Fetched data for", user_id);
        return user;
    }
    catch (error) {
        console.log("[U] Failed data fetch for user.");
        return false;
    }
}

// Intune deletion
export async function deleteIntuneDevices(DEVICES) {
    for (const device of DEVICES) {
        console.log("[I] Deleting device: " + device);
        
        try {
            const deletion = await _appClient
            .api(`/deviceManagement/managedDevices/${device}`)
            //.get();
            .delete();

            console.log("[I] Device deleted - " + deletion);
            return true;
        } catch (error) {
            // Error
            console.log(error);
            return false;
        }
    };

    // Check that device is deleted in the interval of one minute. Max checks 15, before returning false.
    const interval = 60000, maxChecks = 20;
    for (const device of DEVICES){
        let deleted = false;
        let attempts = 0;
        
        while (!deleted && attempts < maxChecks) {
            console.log("[I] Autopilot deletion request(s) sent. Checking if device(s) exist (attempts left", (maxChecks - attempts) + ")");
            try {
                await new Promise(resolve => setTimeout(resolve, interval));
                const check = await _appClient
                .api(`/deviceManagement/managedDevices/${device}`)
                .get();
            } catch (error) {
                if (error.statusCode === 404) {
                    console.log(`[I] Device ${device} deletion confirmed at ${attempts} attempts.`)
                    deleted = true;
                } else {
                    console.log("[I] API request returned a non-404 error:", error);
                }
            }
            attempts++;
        }

        if (!deleted) {
            console.log("[I] Max attempts reached for", device + ". Device deletion not confirmed.");
            return false;
        }
    }

    console.log("[I] Continuing to Autopilot");
    return true;
}

// Autopilot deletion
export async function deleteAutopilotDevices(DEVICES) {
    for (const device of DEVICES) {
        console.log("[A] Deleting device: " + device);
        try {
            const deletion = await _appClient
            .api(`/deviceManagement/windowsAutopilotDeviceIdentities/${device}`)
            //.get();
            .delete();

            console.log("[A] Device deleted - " + deletion);
        } catch (error) {
            console.log(error);
            return false;
        }
    };

    const interval = 60000, maxChecks = 20;
    for (const device of DEVICES){
        let deleted = false;
        let attempts = 0;
        
        while (!deleted && attempts < maxChecks) {
            console.log("[A] Autopilot deletion request(s) sent. Checking if device(s) exist (attempts left", (maxChecks - attempts) + ")");
            try {
                await new Promise(resolve => setTimeout(resolve, interval));
                const check = await _appClient
                .api(`/deviceManagement/windowsAutopilotDeviceIdentities/${device}`)
                .get();
            } catch (error) {
                if (error.statusCode === 404) {
                    console.log(`[A] Device ${device} deletion confirmed.`)
                    deleted = true;
                } else {
                    console.log("[A] API request returned a non-404 error:", error);
                }
            }
            attempts++;
        }

        if (!deleted) {
            console.log("[A] Max attempts reached for", device + ". Device deletion not confirmed.");
            return false;
        }
    }

    console.log("[A] Continuing to Entra");
    return true;
}

// Entra deletion
export async function deleteEntraDevices(DEVICES) {
    for (const device of DEVICES) {
        console.log("[E] Deleting device: " + device);
        try {
            const deletion = await _appClient
            .api(`/devices/${device}`)
            //.get();
            .delete();

            console.log("[E] Device deleted - " + device);
        } catch (error) {
            console.log(error);
            return false;
        }
    };

    const interval = 60000, maxChecks = 20;
    for (const device of DEVICES){
        let deleted = false;
        let attempts = 0;
        
        while (!deleted && attempts < maxChecks) {
            console.log("[E] Entra deletion request(s) sent. Checking if device(s) exist (attempts left", (maxChecks - attempts) + ")");
            try {
                await new Promise(resolve => setTimeout(resolve, interval));
                const check = await _appClient
                .api(`/devices/${device}`)
                .get();

            } catch (error) {
                if (error.statusCode === 404) {
                    console.log(`[E] Device ${device} deletion confirmed.`)
                    deleted = true;
                } else {
                    console.log("[E] API request returned a non-404 error:", error);
                }
            }
            attempts++;
        }

        if (!deleted) {
            console.log("[E] Max attempts reached for", device + ". Device deletion not confirmed.");
            return false;
        }
    }

    console.log("[E] Deletions complete.");
    return true;
}
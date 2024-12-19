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
    var filtered_serials = serial_array.filter(elm => elm);
    var intune = [], autopilot = [], entra = [], result = [];

    // Each platform has it's own function for data fetching
    intune = await fetchIntune(filtered_serials);
    autopilot = await fetchAutopilot(filtered_serials);
    entra = await fetchEntra(filtered_serials);

    // Merge data from all the arrays into a final version
    for (let i = 0; i < (filtered_serials.length); i++) {
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
    let allDeleted = true;
    for (const device of DEVICES) {
        console.log("[I] Deleting device: " + device);
        
        try {
            const deletion = await _appClient
            .api(`/deviceManagement/managedDevices/${device}`)
            .delete();
            console.log("[I] Device deleted - " + device);
        } catch (error) {
            // Error
            console.log("[I] Error!:", error);
            return false;
        }
    };

    // Check that device is deleted in the interval of one minute. Max checks 15, before returning false.
    const interval = 60000, maxChecks = 20;
    for (const device of DEVICES){
        let deleted = false;
        let attempts = 0;
        
        while (!deleted && attempts < maxChecks) {
            console.log("[I] Intune deletion request(s) sent. Checking if device(s) exist (attempts left", (maxChecks - attempts) + ")");
            await new Promise(resolve => setTimeout(resolve, interval));
            try {
                const check = await _appClient
                .api(`/deviceManagement/managedDevices/${device}`)
                .get();
                console.log("[I] Device " + device + " still found.");
            } catch (error) {
                if (error.statusCode === 404) {
                    console.log(`[I] Device ${device} deletion confirmed at ${attempts} attempts.`);
                    deleted = true;
                } else {
                    console.log("[I] API request returned a non-404 error:", error);
                }
            }
            attempts++;
        }

        if (!deleted) {
            console.log("[I] Max attempts reached for", device + ". Device deletion not confirmed.");
            allDeleted = false;
        }
    }

    if (allDeleted) {
        console.log("[I] Intune deletions complete.");
        return true;
    } else {
        return false;
    }
}

// Autopilot deletion
export async function deleteAutopilotDevices(DEVICES) {
    let allDeleted = true;
    // Loop all devices and delete them
    for (const device of DEVICES) {
        console.log("[A] Deleting device: " + device);
        try {
            const deletion = await _appClient
                .api(`/deviceManagement/windowsAutopilotDeviceIdentities/${device}`)
                .delete();
            console.log("[A] Device deleted - " + device);
        } catch (error) {
            if (error.statusCode === 400) {
                console.log("[A] Device already queried for deletion, skipping. (400)");
            } else {
                console.log("[A] ABORTING: Error during deletion:", error);
                return false;
            }
        }
    }

    // Check if the devices are deleted or not
    let initialInterval = 30000; // 30 second inverval
    const maxChecks = 6; // 10 attempts before returning false
    
    // Loop through devices, again
    for (const device of DEVICES) {
        let deleted = false;
        let attempts = 0;

        // While deleted flag is false and attempts aren't more than the max. amount
        while (!deleted && attempts < maxChecks) {
            try {
                // Wait for 30 seconds -> 60 seconds -> 120 seconds etc. (max. 16 min @ attempt )
                const waitTime = initialInterval * Math.pow(2, attempts);
                console.log(`[A] Waiting ${waitTime / 1000} seconds before checking deletion status for device ${device} (attempt ${attempts + 1})`);
                await new Promise(resolve => setTimeout(resolve, waitTime));
                // Fetch all devices instead of a single one to bypass caching
                const allDevices = await _appClient
                    .api(`/deviceManagement/windowsAutopilotDeviceIdentities`)
                    .get();
                
                // Check if device is still in the list
                if (!allDevices.value.find(d => d.id === device)) {
                    console.log(`[A] Device ${device} confirmed deleted.`);
                    deleted = true;
                } else {
                    console.log(`[A] Device ${device} still exists.`);
                }
                // We can use 404 to confirm deletion. Other codes, should be noted.
            } catch (error) {
                if (error.statusCode === 404) {
                    console.log(`[A] Device ${device} deletion confirmed by 404.`);
                    deleted = true;
                } else {
                    console.log("[A] API request returned a non-404 error:", error);
                }
            }
            attempts++;
        }
        
        // If deletion doesn't go through in a given time, the program returns false
        if (!deleted) {
            console.log("[A] Max attempts reached for", device + ". Device deletion not confirmed.");
            allDeleted = false;
        }
    }

    // If all devices have been confirmed to be deleted, continue
    if (allDeleted) {
        console.log("[A] Autopilot deletions complete.");
        return true;
    } else {
        return false;
    }
}

// Entra deletion
export async function deleteEntraDevices(DEVICES) {
    let allDeleted = true;
    // Loop through devices and delete them
    for (const device of DEVICES) {
        console.log("[E] Deleting device: " + device);
        try {
            const deletion = await _appClient
            .api(`/devices/${device}`)
            .delete();

            console.log("[E] Device deleted - " + device);
        } catch (error) {
            if (error.statusCode === 404) {
                console.log("[E] Device not found or is already deleted. (404)");
            } else {
                console.log("[E] ABORTING: Error while deleting device:", error);
                return false;
            }
        }
    };

    // In Entra, we don't have to worry about cache that much so we just set 20 one minute checks
    const interval = 60000, maxChecks = 20;
    for (const device of DEVICES){
        let deleted = false;
        let attempts = 0;
        
        while (!deleted && attempts < maxChecks) {
            console.log("[E] Entra deletion request(s) sent. Checking if device(s) exist (attempts left", (maxChecks - attempts) + ")");
            await new Promise(resolve => setTimeout(resolve, interval));
            
            try {
                const check = await _appClient
                .api(`/devices/${device}`)
                .get();
                console.log("Returned check:", check);
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
            allDeleted = false;
        }
    }

    if (allDeleted) {
        console.log("[E] Entra deletions complete.");
        return true;
    } else {
        return false;
    }
}
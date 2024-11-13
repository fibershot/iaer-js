// In interaction with index.html
// Function activated from "Submit serials" button.
var public_data;
async function submitSerials () {
    console.log("Request received");
    // Fetch data from input field
    const input = document.getElementById("serial").elements[0].value;

    if (!input) {
        console.log("Invalid input! =>", input);
        window.alert("Serial field cannot be empty.");
    } else {
        const resultDiv = document.getElementById("resultDiv");
        const logDiv = document.getElementById("log");
        const loading = document.createElement("img");
        loading.src="/media/spinner.gif"; loading.style="padding-left: 1.5em; max-width: 50px; max-height: 50px;";
        resultDiv.innerHTML = "<br>&nbsp;<br>Fetching data...";
        resultDiv.appendChild(loading);
        // Send request for data
        try {
            const response = await fetch("/api/fetch-data", {
                method: "POST",
                headers: {
                    "Content-Type": "application/json",
                },
                body: JSON.stringify({ serials: input })
            });
    
            const result = await response.json();
    
            for (let i = 0; i < result.results.length; i++) {
                console.log("Found:", result.results[i]);
            }
    
            // If data returned was OK (200)
            if (response.ok) {
                // Get resultDiv and empty it
                resultDiv.innerHTML = "";
                logDiv.textContent = "";

                // Create header
                const headerRow = document.createElement("div");
                headerRow.className = "device-row header";
                const headers = ["Name", "Intune", "Autopilot", "Entra", "Select"];
                headers.forEach(text => {
                    const headerSpan = document.createElement("span");
                    headerSpan.textContent = text;
                    headerRow.appendChild(headerSpan);
                });
            
                resultDiv.appendChild(headerRow);
            
                // variable, data, is the returned data we got from the fetch request
                const data = result.results;
                public_data = result.results;
                // Add data to rows
                for (let i = 0; i < result.results.length; i++) {

                    if (!data[i].intuneFound && !data[i].autopilotFound && !data[i].intuneFound){
                        console.log("The search with " + data[i].initialReq + " returned no results.");
                        logDiv.textContent += `${data[i].initialReq} invalid.\r\n`;
                        continue;
                    }

                    const deviceRow = document.createElement("div");
                    deviceRow.className = "device-row";
            
                    // Device name
                    const nameSpan = document.createElement("span");
                    nameSpan.className = "device-name";
                    nameSpan.onclick = function() {
                        toggleDetails(nameSpan);
                    }

                    if (!data[i].intuneFound) {
                        nameSpan.textContent = data[i].autopilotName;
                        if (!data[i].autopilotFound) {
                            nameSpan.textContent = data[i].entraName;
                        }
                    } else { nameSpan.textContent = data[i].intuneName; }
                    
                    deviceRow.appendChild(nameSpan);
            
                    // Intune, Autopilot, and Entra check
                    const intuneSpan = document.createElement("span");
                    intuneSpan.textContent = data[i].intuneFound;
                    intuneSpan.style.color = checkColor(data[i].intuneFound);
                    deviceRow.appendChild(intuneSpan);
            
                    const autopilotSpan = document.createElement("span");
                    autopilotSpan.textContent = data[i].autopilotFound;
                    autopilotSpan.style.color = checkColor(data[i].autopilotFound);
                    deviceRow.appendChild(autopilotSpan);
            
                    const entraSpan = document.createElement("span");
                    entraSpan.textContent = data[i].entraFound;
                    entraSpan.style.color = checkColor(data[i].entraFound);
                    deviceRow.appendChild(entraSpan);
            
                    // Select checkbox
                    const checkBoxSpan = document.createElement("span");
                    const checkBox = document.createElement("input");
                    checkBox.type = "checkbox";
                    checkBox.className = "markedForDeletion";
                    checkBox.id = data[i].intuneResult + "-" + data[i].autopilotResult + "-" + data[i].entraResult;
                    checkBoxSpan.appendChild(checkBox);
                    deviceRow.appendChild(checkBoxSpan);
            
                    // Add more info from interacting with the device name
                    const detailDiv = document.createElement("div");
                    detailDiv.className = "device-details";
                    const detailData = document.createElement("p");

                    const options = { 
                        year: 'numeric', 
                        month: 'numeric', 
                        day: 'numeric', 
                        hour: 'numeric', 
                        minute: 'numeric', 
                        second: 'numeric' 
                    };
                    
                    detailData.innerHTML = `
                        <strong>Login</strong><br>
                        Last login user: ${data[i].lastLogOnUser} <br>
                        Last login time: ${new Date(data[i].lastLogin).toLocaleDateString(undefined, options)} <br>
                        Last contact time: ${new Date(data[i].lastContactedDateTime).toLocaleDateString(undefined, options)}<br>
                        Last login email: ${data[i].lastLogOnUserEmail} <br>
                        User principal name: ${data[i].userPrincipalName}<br><br>

                        <strong>Intune</strong><br>
                        Intune found: ${data[i].intuneFound} <br>
                        Intune ID: ${data[i].intuneId} <br>
                        Intune name: ${data[i].intuneName} <br>
                        Intune serial: ${data[i].intuneSerial}<br><br>

                        <strong>Autopilot</strong><br>
                        Autopilot found: ${data[i].autopilotFound} <br>
                        Autopilot ID: ${data[i].autopilotId} <br>
                        Autopilot name: ${data[i].autopilotName} <br>
                        Autopilot serial: ${data[i].autopilotSerial}<br><br>
                        
                        <strong>Entra</strong><br>
                        Entra found: ${data[i].entraFound} <br>
                        Entra ID: ${data[i].entraId} <br>
                        Entra name: ${data[i].entraName} <br><br>

                        <strong>Hardware</strong><br>
                        Model: ${data[i].model}<br><br>
                    `;
                    detailDiv.appendChild(detailData);
                    resultDiv.appendChild(deviceRow);
                    resultDiv.appendChild(detailDiv);
                    
                }
            }
    
        } catch (error) {
            console.log("Nothing returned." + error);
        }
    }
}

// Change color for the Intune, Autopilot and Entra section
function checkColor(object){
    if (object){
        return "green";
    } else {
        return "red";
    }
}

// Show more information
function toggleDetails(element) {
    var detailsDiv = element.parentNode.nextElementSibling;
    detailsDiv.classList.toggle('show');
}

async function deleteSerials() {

    var selected = [];
    var check_consistency, removeID;
    var checkboxes = document.querySelectorAll(".markedForDeletion");
    var intuneID, autopilotID, entraID;

    checkboxes.forEach((checkbox) => {
        if (checkbox.checked) {
            selected.push(checkbox.id);
            console.log(checkbox.id);

            // Want to make sure we've fetched the correct values
            check_consistency = checkbox.id.split("-");
            if (check_consistency[0] === check_consistency[1] && check_consistency[1] === check_consistency[2]){
                console.log("Correct value in all checks, continuing deletion with ID", check_consistency[0], "in public_data");
                removeID = check_consistency[0];

                // Send device for removal :(
                // Jokes aside. Use platform IDs for deletion.
                if (public_data[removeID].intuneFound) {
                    intuneID = public_data[removeID].intuneId;
                }
                if (public_data[removeID].autopilotFound) {
                    autopilotID = public_data[removeID].autopilotId;
                }
                if (public_data[removeID].entraFound) {
                    entraID = public_data[removeID].entraId;
                }

                console.log("Intune ID:", intuneID, "\nAutopilot ID:", autopilotID, "\nEntra ID:", entraID);

            } else {
                console.error("IDs do NOT match. Fetch cleared, but is unstable for use!");
            }
        }
    });
}
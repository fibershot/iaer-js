async function submitSerials () {
    console.log("Request received");
    const input = document.getElementById("serial").elements[0].value;

    if (!input) {
        console.log("Invalid input! =>", input);
        window.alert("Serial field cannot be empty.");
    } else {
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
    
            if (response.ok) {
                const resultDiv = document.getElementById("resultDiv");
                resultDiv.innerHTML = "";
            
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
            
                const data = result.results;
                // Add data to rows
                for (let i = 0; i < result.results.length; i++) {
                    const deviceRow = document.createElement("div");
                    deviceRow.className = "device-row";
            
                    // Device name
                    const nameSpan = document.createElement("span");
                    nameSpan.className = "device-name";
                    nameSpan.onclick = function() {
                        toggleDetails(nameSpan);
                    }
                    nameSpan.textContent = data[i].intuneName;
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
            
                    // Remove checkbox
                    const checkBoxSpan = document.createElement("span");
                    const checkBox = document.createElement("input");
                    checkBox.type = "checkbox";
                    checkBoxSpan.appendChild(checkBox);
                    deviceRow.appendChild(checkBoxSpan);
            
                    // Add everything
                    const detailDiv = document.createElement("div");
                    detailDiv.className = "device-details";
                    const detailData = document.createElement("p");
                    detailData.innerHTML = `
                        Last login user: ${data[i].lastLogOnUser} <br>
                        Last login time: ${data[i].lastLogin} <br>
                        Last contact time: ${data[i].lastContactedDateTime}<br>
                        Last login email: ${data[i].lastLogOnUserEmail} <br>
                        User principal name: ${data[i].userPrincipalName}<br><br>

                        Model: ${data[i].model}<br><br>

                        Autopilot found: ${data[i].autopilotFound} <br>
                        Autopilot ID: ${data[i].autopilotId} <br>
                        Autopilot name: ${data[i].autopilotName} <br><br>
                        
                        Entra found: ${data[i].entraFound} <br>
                        Entra ID: ${data[i].entraId} <br>
                        Entra name: ${data[i].entraName} <br><br>
                        
                        Intune found: ${data[i].intuneFound} <br>
                        Intune ID: ${data[i].intuneId} <br>
                        Intune name: ${data[i].intuneName} <br><br>
                    `;
                    detailDiv.appendChild(detailData);

                    resultDiv.appendChild(deviceRow);
                    resultDiv.appendChild(detailDiv);
                    
                }
            }
    
        } catch (error) {
            console.log("Nothing returned.");
        }
    }
}

function checkColor(object){
    if (object){
        return "green";
    } else {
        return "red";
    }
}

function toggleDetails(element) {
    var detailsDiv = element.parentNode.nextElementSibling;
    detailsDiv.classList.toggle('show');
}
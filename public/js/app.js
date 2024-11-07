async function submitSerials () {
    console.log("Request received");
    const input = document.getElementById("serial").elements[0].value;

    if (!input) {
        console.log("Invalid input! =>", input);
    }

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
            console.log("Found:", result.results[i].deviceName);
        }

        if (response.ok) {
            const resultDiv = document.getElementById("resultDiv");
            resultDiv.innerHTML = "";
        
            // Create header
            const headerRow = document.createElement("div");
            headerRow.className = "device-row header";
        
            const headers = ["Device name", "Intune", "Autopilot", "Entra", "Remove"];
            headers.forEach(text => {
                const headerSpan = document.createElement("span");
                headerSpan.textContent = text;
                headerRow.appendChild(headerSpan);
            });
        
            resultDiv.appendChild(headerRow);
        
            // Add data to rows
            for (let i = 0; i < result.results.length; i++) {
                const deviceRow = document.createElement("div");
                deviceRow.className = "device-row";
        
                // Device name
                const nameSpan = document.createElement("span");
                nameSpan.textContent = result.results[i].deviceName;
                deviceRow.appendChild(nameSpan);
        
                // Intune, Autopilot, and Entra check
                const intuneSpan = document.createElement("span");
                intuneSpan.textContent = "X";
                deviceRow.appendChild(intuneSpan);
        
                const autopilotSpan = document.createElement("span");
                autopilotSpan.textContent = "X";
                deviceRow.appendChild(autopilotSpan);
        
                const entraSpan = document.createElement("span");
                entraSpan.textContent = "X";
                deviceRow.appendChild(entraSpan);
        
                // Remove checkbox
                const checkBoxSpan = document.createElement("span");
                const checkBox = document.createElement("input");
                checkBox.type = "checkbox";
                checkBox.checked = true;
                checkBoxSpan.appendChild(checkBox);
                deviceRow.appendChild(checkBoxSpan);
        
                // Add everything
                resultDiv.appendChild(deviceRow);
            }
        }

    } catch (error) {
        console.log(error);
    }
}
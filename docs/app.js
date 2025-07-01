// Function to show notification
function showNotification(message, isError = false) {
    // Create notification element if it doesn't exist
    let notification = document.getElementById('notification');
    if (!notification) {
        notification = document.createElement('div');
        notification.id = 'notification';
        notification.style.cssText = `
            position: fixed;
            top: 20px;
            right: 20px;
            padding: 15px 25px;
            border-radius: 4px;
            color: white;
            font-weight: 500;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
            z-index: 1000;
            transform: translateX(120%);
            transition: transform 0.3s ease-in-out;
            max-width: 350px;
        `;
        document.body.appendChild(notification);
    }

    // Set notification content and style
    notification.textContent = message;
    notification.style.backgroundColor = isError ? '#f44336' : '#4CAF50';
    notification.style.transform = 'translateX(0)';

    // Auto-hide after 5 seconds
    setTimeout(() => {
        notification.style.transform = 'translateX(120%)';
    }, 5000);
}

// Function to create agent by calling the Logic App
async function createAgent() {
    const agentName = document.getElementById('agentName').value.trim();
    const model = document.getElementById('modelSelect').value;
    
    if (!agentName) {
        showNotification('Please enter a name for your agent', true);
        return;
    }
    
    const createAgentBtn = document.getElementById('createAgentBtn');
    const originalText = createAgentBtn.textContent;
    
    try {
        // Disable button and show loading state
        createAgentBtn.disabled = true;
        createAgentBtn.textContent = 'Creating...';
        
        // Show creating notification
        showNotification(`Creating agent "${agentName}" with model ${model}...`);
        
        // Call the Logic App to create the agent
        const url = "https://prod-41.westus.logic.azure.com:443/workflows/e5f0ce23f3ea415696da0d9b4eeed2ec/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=IZXxoQiXyN8FToQ0GSaFPAy8iO9NEDf9vx5qRP7g0NA";
        
        const requestBody = {
            botName: agentName,
            botModel: model,
            timestamp: new Date().toISOString()
        };
        
        const response = await fetch(url, {
            method: "POST",
            headers: {
                "Content-Type": "application/json"
            },
            body: JSON.stringify(requestBody)
        });
        
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        
        const data = await response.json();
        console.log("Flow response:", data);
        
        // Show success notification
        showNotification(`✅ Agent "${agentName}" created successfully!`);
        
        // Reset form
        document.getElementById('agentName').value = '';
        
    } catch (error) {
        console.error("Error creating agent:", error);
        showNotification(`❌ Failed to create agent: ${error.message}`, true);
    } finally {
        // Re-enable button and restore text
        createAgentBtn.disabled = false;
        createAgentBtn.textContent = originalText;
    }
}

// Function to show notification
function showNotification(message, isError = false) {
    // Create notification element if it doesn't exist
    let notification = document.getElementById('notification');
    if (!notification) {
        notification = document.createElement('div');
        notification.id = 'notification';
        notification.style.cssText = `
            position: fixed;
            top: 20px;
            right: 20px;
            padding: 15px 25px;
            border-radius: 4px;
            color: white;
            font-weight: 500;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
            z-index: 1000;
            transform: translateX(120%);
            transition: transform 0.3s ease-in-out;
            max-width: 350px;
        `;
        document.body.appendChild(notification);
    }

    // Set notification content and style
    notification.textContent = message;
    notification.style.backgroundColor = isError ? '#f44336' : '#4CAF50';
    notification.style.transform = 'translateX(0)';

    // Auto-hide after 5 seconds
    setTimeout(() => {
        notification.style.transform = 'translateX(120%)';
    }, 5000);
}

// Initialize the application when the page loads
document.addEventListener('DOMContentLoaded', function() {
    // Add click event for create agent button
    const createAgentBtn = document.getElementById('createAgentBtn');
    createAgentBtn.addEventListener('click', createAgent);

    //My Code


    // Initialize Microsoft Teams SDK
microsoftTeams.app.initialize().then(() => {
    microsoftTeams.app.getContext().then((context) => {
        // Log team and channel info
        console.log("Team Name:", context.team?.displayName || "Not available");
        console.log("Channel Name:", context.channel?.displayName || "Not available");
        console.log("Channel Type:", context.channel?.membershipType || "Unknown");

        // Build URLs using your domain
        let standardUrl = "Not available";
        let privateUrl = "Not available";

        if (context.sharePointSite?.teamSiteUrl) {
            // Standard channel site
            standardUrl = `${context.sharePointSite.teamSiteUrl}/Shared%20Documents`;

            // Private channel site (private channels have a separate site URL)
            if (context.channel?.membershipType === "Private" && context.sharePointSite.privateSiteUrl) {
                privateUrl = `${context.sharePointSite.privateSiteUrl}/Shared%20Documents`;
            }
        }

        // Log URLs
        console.log("Standard Channel URL:", standardUrl);
        console.log("Private Channel URL:", privateUrl);
    });
});



    showNotification('✅ App loaded successfully!'+privateUrl);
    // Hide login button if it exists
    const loginBtn = document.getElementById('loginBtn');
    if (loginBtn) {
        loginBtn.style.display = 'none';
    }
    
    
    // Add animation to the button on hover
    createAgentBtn.addEventListener('mouseenter', function() {
        this.style.transform = 'translateY(-2px)';
        this.style.boxShadow = '0 6px 16px rgba(121, 80, 242, 0.2)';
    });

    createAgentBtn.addEventListener('mouseleave', function() {
        this.style.transform = 'translateY(0)';
        this.style.boxShadow = 'none';
    });
});

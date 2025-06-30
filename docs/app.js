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
            teamsContext: {
                teamId: teamsContext.teamId,
                teamName: teamsContext.teamName,
                channelId: teamsContext.channelId,
                channelName: teamsContext.channelName,
                sharePointSiteUrl: teamsContext.sharePointSiteUrl,
                documentLibraryUrl: getDocumentLibraryUrl(teamsContext.sharePointSiteUrl)
            },
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

// Initialize Microsoft Teams
microsoftTeams.initialize();

// Store Teams context
let teamsContext = {
    teamId: '',
    channelId: '',
    channelName: '',
    teamName: '',
    sharePointSiteUrl: ''
};

// Function to get Teams context
async function getTeamsContext() {
    return new Promise((resolve, reject) => {
        microsoftTeams.getContext(context => {
            const isPrivate = getChannelType(context.channelId) === 'private';
            const channelName = context.channelName || 'General';
            const teamName = context.teamName || 'Team';
            
            // For private channels, we need to handle the SharePoint URL differently
            let sharePointSiteUrl = context.teamSiteUrl || '';
            
            teamsContext = {
                teamId: context.teamId,
                channelId: context.channelId,
                channelName: channelName,
                teamName: teamName,
                sharePointSiteUrl: sharePointSiteUrl,
                isPrivate: isPrivate,
                channelType: isPrivate ? 'private' : 'standard'
            };
            
            // If this is a private channel, we need to get the SharePoint site URL
            if (isPrivate) {
                // For private channels, the site URL is different
                // We'll use the teamSiteUrl as a base and modify it
                teamsContext.sharePointSiteUrl = sharePointSiteUrl.replace(
                    /\/sites\/([^/]+)/,
                    `/sites/${teamName.replace(/\s+/g, '')}-${channelName.replace(/\s+/g, '')}`
                );
            }
            
            resolve(teamsContext);
        });
    });
}

// Function to get SharePoint document library URL based on channel type
function getDocumentLibraryUrl(siteUrl, channelName, isPrivate) {
    if (!siteUrl) return '';
    
    // For standard channels, the URL is straightforward
    if (!isPrivate) {
        return `${siteUrl}/Shared%20Documents`;
    }
    
    // For private channels, the URL follows a different pattern
    // Example: https://contoso.sharepoint.com/sites/TeamName-PrivateChannelName-{GUID}
    // We need to get the actual site URL from the context
    return `${siteUrl}/Shared%20Documents`;
}

// Function to get channel type (standard or private)
function getChannelType(channelId) {
    // In Teams, private channels have a specific GUID format
    // This is a simplified check - in production, you might want to use the Teams API
    return channelId && channelId.startsWith('19:') ? 'private' : 'standard';
}

// Initialize the application when the page loads
document.addEventListener('DOMContentLoaded', async function() {
    // Initialize Teams context
    try {
        await getTeamsContext();
        
        // Update UI with Teams context
        const teamInfo = document.getElementById('teamInfo');
        if (teamInfo) {
            const docLibUrl = getDocumentLibraryUrl(
                teamsContext.sharePointSiteUrl, 
                teamsContext.channelName, 
                teamsContext.isPrivate
            );
            
            teamInfo.innerHTML = `
                <div class="info-group">
                    <div class="info-label">Team:</div>
                    <div class="info-value">${teamsContext.teamName}</div>
                </div>
                <div class="info-group">
                    <div class="info-label">Channel:</div>
                    <div class="info-value">
                        ${teamsContext.channelName}
                        <span style="font-size: 0.8em; color: ${teamsContext.isPrivate ? '#d83b01' : '#107c10'}; margin-left: 8px;">
                            (${teamsContext.isPrivate ? 'Private' : 'Standard'} Channel)
                        </span>
                    </div>
                </div>
                <div class="info-group">
                    <div class="info-label">SharePoint Site:</div>
                    <div class="info-value">
                        <a href="${teamsContext.sharePointSiteUrl}" target="_blank" style="word-break: break-all;">
                            ${teamsContext.sharePointSiteUrl}
                        </a>
                    </div>
                </div>
                ${docLibUrl ? `
                <div class="info-group">
                    <div class="info-label">Document Library:</div>
                    <div class="info-value">
                        <a href="${docLibUrl}" target="_blank">Open in SharePoint</a>
                    </div>
                </div>` : ''}
            `;
        }
    } catch (error) {
        console.error('Error initializing Teams context:', error);
    }

    // Add click event for create agent button
    document.getElementById('createAgentBtn').addEventListener('click', createAgent);
    
    // Hide login button as we're using Teams authentication
    document.getElementById('loginBtn').style.display = 'none';
    
    // Add animation to the button on hover
    const createAgentBtn = document.getElementById('createAgentBtn');
    createAgentBtn.addEventListener('mouseenter', function() {
        this.style.transform = 'translateY(-2px)';
        this.style.boxShadow = '0 6px 16px rgba(121, 80, 242, 0.2)';
    });

    createAgentBtn.addEventListener('mouseleave', function() {
        this.style.transform = 'translateY(0)';
        this.style.boxShadow = 'none';
    });
});

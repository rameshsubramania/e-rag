
// Debug logging function
function logDebug(message, data = null) {
    const timestamp = new Date().toISOString();
    const logMessage = `[${timestamp}] ${message}`;
    console.log(logMessage);
    if (data) {
        console.log('Data:', JSON.stringify(data, null, 2));
    }

    // Also show in UI if debug element exists
    const debugElement = document.getElementById('debug');
    if (debugElement) {
        const debugLine = document.createElement('div');
        debugLine.textContent = logMessage;
        if (data) {
            const dataLine = document.createElement('div');
            dataLine.textContent = JSON.stringify(data, null, 2);
            dataLine.style.marginLeft = '20px';
            debugElement.appendChild(dataLine);
        }
        debugElement.appendChild(debugLine);
    }
}

// Error logging function
function logError(message, error = null) {
    const timestamp = new Date().toISOString();
    const errorMessage = `[${timestamp}] ERROR: ${message}`;
    console.error(errorMessage);
    if (error) {
        console.error('Error details:', error);
    }

    // Also show in UI if debug element exists
    const debugElement = document.getElementById('debug');
    if (debugElement) {
        const errorLine = document.createElement('div');
        errorLine.style.color = 'red';
        errorLine.textContent = errorMessage;
        if (error) {
            const errorDetails = document.createElement('div');
            errorDetails.style.color = 'red';
            errorDetails.style.marginLeft = '20px';
            errorDetails.textContent = JSON.stringify(error, null, 2);
            debugElement.appendChild(errorDetails);
        }
        debugElement.appendChild(errorLine);
    }
}

// Sanitize string for URL
function sanitizeForUrl(str) {
    return str
        .toLowerCase()
        .replace(/[^a-z0-9]+/g, '-')
        .replace(/^-+|-+$/g, '');
}

// Copy URL function
function copyUrl() {
    const urlElement = document.getElementById('sharepointUrl');
    const url = urlElement.textContent;
    
    navigator.clipboard.writeText(url).then(() => {
        const successMsg = document.getElementById('copySuccess');
        successMsg.style.display = 'block';
        setTimeout(() => {
            successMsg.style.display = 'none';
        }, 3000);
    }).catch(err => {
        logError('Failed to copy URL', err);
    });
}

// Main initialization function
function initializeTeamsApp() {
    try {
        logDebug('Initializing Teams app...');
        microsoftTeams.app.initialize().then(() => {
            logDebug('Teams app initialized successfully');
            return microsoftTeams.app.getContext();
        }).then((context) => {
            logDebug('Got Teams context:', context);

            if (!context) {
                throw new Error('Teams context is null or undefined');
            }

            logDebug('Teams context received', context);

            // Extract tenant ID with validation
            const tenantId = context.tid || context.user?.tenant?.id || '';
            logDebug('Tenant ID:', tenantId);

            // Extract tenant name from user's principal name
            const tenantName = context.user?.userPrincipalName?.split('@')[1]?.split('.')[0] || '';
            logDebug('Tenant Name:', tenantName);
            document.getElementById('tenantName').textContent = tenantName || 'Not available';

            // Extract team info
            const teamId = context.team?.internalId || context.team?.id || 'Not available';
            const teamName = context.team?.displayName || 'Not available';
            logDebug('Team Info:', { id: teamId, name: teamName });

            // Extract channel info
            const channelId = context.channel?.id || 'Not available';
            const channelName = context.channel?.displayName || 'Not available';
            const channelType = context.channel?.membershipType || 'Not available';
            logDebug('Channel Info:', { id: channelId, name: channelName, type: channelType });

            // Update UI with values
            document.getElementById('teamId').textContent = teamId;
            document.getElementById('teamName').textContent = teamName;
            document.getElementById('channelId').textContent = channelId;
            document.getElementById('channelName').textContent = channelName;
            document.getElementById('channelType').textContent = channelType;

            // --- Store the generated SharePoint URL here ---
            let generatedSharepointUrl = ''; // Initialize variable

            if (teamName !== 'Not available' && channelName !== 'Not available') {
                console.log('SharePoint site info:', JSON.stringify(context.sharePointSite, null, 2));

                if (context.channel?.membershipType === 'Private') {
                    generatedSharepointUrl = context.sharePointSite.teamSiteUrl + '/Shared%20Documents';
                    console.log('Using private channel URL:', generatedSharepointUrl);
                } else {
                    generatedSharepointUrl = context.sharePointSite.teamSiteUrl + '/Shared%20Documents';
                    console.log('Using standard channel URL:', generatedSharepointUrl);
                }
                document.getElementById('sharepointUrl').textContent = generatedSharepointUrl; // Update display
            } else {
                document.getElementById('sharepointUrl').textContent = 'Cannot generate URL - missing team or channel name';
                generatedSharepointUrl = 'N/A'; // Set to N/A if cannot generate
            }
            // --- End SharePoint URL Generation ---


            // ***** MODIFIED PART FOR POWER APPS EMBEDDING *****
            const powerAppIframe = document.getElementById('powerAppIframe');
            const powerAppId = '3243308d-d91c-4948-a5e3-e98e3a7d8ae5'; // Your specific Power App ID

            if (powerAppIframe && teamId !== 'Not available' && channelId !== 'Not available') {
                let powerAppUrl = `https://apps.powerapps.com/play/${powerAppId}?source=website`;

                powerAppUrl += `&sharepointUrl=${encodeURIComponent(generatedSharepointUrl)}`; // Parameter 1
                powerAppUrl += `&channelId=${encodeURIComponent(channelId)}`; // Parameter 3
                powerAppUrl += `&channelName=${encodeURIComponent(channelName)}`; // Parameter 4
                powerAppUrl += `&tenantId=${encodeURIComponent(tenantId)}`; // Parameter 5
                powerAppUrl += `&teamId=${encodeURIComponent(teamId)}`; // Parameter 6
                powerAppUrl += `&teamName=${encodeURIComponent(teamName)}`; // Parameter 7
                powerAppUrl += `&channelType=${encodeURIComponent(channelType)}`; // Parameter 8

                // --- ADD THE SHAREPOINT URL PARAMETER HERE --- // This is a comment, not code
                console.log('Power App Embed URL set:', powerAppUrl); // Log 1
                console.log('Debug: Value of generatedSharepointUrl BEFORE append:', generatedSharepointUrl);

                // FIRST ASSIGNMENT TO IFRAME SRC
                try {
                    logDebug('Setting Power App iframe URL', powerAppUrl);
                    powerAppIframe.src = powerAppUrl;
                    logDebug('Power App iframe URL set successfully');
                } catch (error) {
                    logError('Failed to set Power App iframe URL', error);
                }
            } else {
                console.warn('Could not embed Power App: Missing iframe element or essential Teams context. Power App will not load with parameters.');
            }
    }).catch(error => {
        logError('Teams initialization or context error', error);
        document.getElementById('debug').textContent = 'Error: Failed to initialize Teams or get context';
    });
    } catch (error) {
        logError('Teams app initialization failed', error);
        document.getElementById('debug').textContent = 'Error: Failed to initialize Teams app. Please refresh and try again.';
    }
}

// Start the app
try {
    logDebug('Starting Teams app...');
    initializeTeamsApp();
} catch (error) {
    logError('Teams app initialization failed', error);
    document.getElementById('debug').textContent = 'Error: Failed to initialize Teams app. Please refresh and try again.';
}

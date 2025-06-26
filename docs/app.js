// // Initialize Teams SDK
// microsoftTeams.app.initialize().then(() => {
//     // Get context
//     microsoftTeams.app.getContext().then((context) => {
//         // Debug information
//         console.log('Teams Context:', JSON.stringify(context, null, 2));
//         // Extract tenant name from user's principal name
//         const tenantName = context.user?.userPrincipalName?.split('@')[1]?.split('.')[0] || '';
//         document.getElementById('tenantName').textContent = tenantName || 'Not available';
//         const teamId = context.team?.internalId || 'Not available';
//         const teamName = context.team?.displayName || 'Not available';
//         const channelId = context.channel?.id || 'Not available';
//         const channelName = context.channel?.displayName || 'Not available';

//         // Update UI with values
//         document.getElementById('teamId').textContent = teamId;
//         document.getElementById('teamName').textContent = teamName;
//         document.getElementById('channelId').textContent = channelId;
//         document.getElementById('channelName').textContent = channelName;

//         // Generate SharePoint URL
//         if (teamName !== 'Not available' && channelName !== 'Not available') {
//             const sanitizedTeamName = sanitizeForUrl(teamName);
//             const sanitizedChannelName = sanitizeForUrl(channelName);
            
//             console.log('SharePoint site info:', JSON.stringify(context.sharePointSite, null, 2));

//             // Use the SharePoint URLs directly from Teams context
//             if (context.channel?.membershipType === 'Private') {
//                 const sharepointUrl = context.sharePointSite.teamSiteUrl + '/Shared%20Documents';
//                 console.log('Using private channel URL:', sharepointUrl);
//                 document.getElementById('sharepointUrl').textContent = sharepointUrl;
//                 document.getElementById('channelType').textContent = 'Private Channel';
//             } else {
//                 // For standard channels, use the relative URL provided by Teams
//                 const sharepointUrl = context.sharePointSite.teamSiteUrl + '/Shared%20Documents';
//                 console.log('Using standard channel URL:', sharepointUrl);
//                 document.getElementById('sharepointUrl').textContent = sharepointUrl;
//                 document.getElementById('channelType').textContent = 'Standard Channel';
//             }
//         } else {
//             document.getElementById('sharepointUrl').textContent = 'Cannot generate URL - missing team or channel name';
//         }
//     });
// });

// // Sanitize string for URL
// function sanitizeForUrl(str) {
//     return str
//         .toLowerCase()
//         .replace(/[^a-z0-9]+/g, '-') // Replace special chars with hyphens
//         .replace(/^-+|-+$/g, ''); // Remove leading/trailing hyphens
// }

// // Copy URL function
// function copyUrl() {
//     const urlElement = document.getElementById('sharepointUrl');
//     const url = urlElement.textContent;
    
//     navigator.clipboard.writeText(url).then(() => {
//         const successMsg = document.getElementById('copySuccess');
//         successMsg.style.display = 'block';
//         setTimeout(() => {
//             successMsg.style.display = 'none';
//         }, 3000);
//     }).catch(err => {
//         console.error('Failed to copy URL:', err);
//     });
// }

//         const teamId = context.team?.internalId || 'Not available';
//         const teamName = context.team?.displayName || 'Not available';
//         const channelId = context.channel?.id || 'Not available';
//         const channelName = context.channel?.displayName || 'Not available';
//         // Use membershipType as it clearly indicates 'Private' or 'Standard'
//         const channelType = context.channel?.membershipType || 'Not available'; 

//         // Update UI with values
//         document.getElementById('teamId').textContent = teamId;
//         document.getElementById('teamName').textContent = teamName;
//         document.getElementById('channelId').textContent = channelId;
//         document.getElementById('channelName').textContent = channelName;
//         document.getElementById('channelType').textContent = channelType;


//         // Generate SharePoint URL (your existing logic)
//         if (teamName !== 'Not available' && channelName !== 'Not available') {
//             // Note: sanitizeForUrl might not be strictly necessary for SharePoint URLs if using context.sharePointSite.teamSiteUrl
//             // but keep it if you need it for other custom path generation
//             // const sanitizedTeamName = sanitizeForUrl(teamName);
//             // const sanitizedChannelName = sanitizeForUrl(channelName);
            
//             console.log('SharePoint site info:', JSON.stringify(context.sharePointSite, null, 2));

//             if (context.channel?.membershipType === 'Private') {
//                 // For private channels, context.sharePointSite.teamSiteUrl points to the private channel's associated site
//                 const sharepointUrl = context.sharePointSite.teamSiteUrl + '/Shared%20Documents';
//                 console.log('Using private channel URL:', sharepointUrl);
//                 document.getElementById('sharepointUrl').textContent = sharepointUrl;
//             } else {
//                 // For standard channels, context.sharePointSite.teamSiteUrl points to the main team site
//                 const sharepointUrl = context.sharePointSite.teamSiteUrl + '/Shared%20Documents';
//                 console.log('Using standard channel URL:', sharepointUrl);
//                 document.getElementById('sharepointUrl').textContent = sharepointUrl;
//             }
//         } else {
//             document.getElementById('sharepointUrl').textContent = 'Cannot generate URL - missing team or channel name';
//         }

//         // ***** THIS IS THE NEW / MODIFIED PART FOR POWER APPS EMBEDDING *****
//         const powerAppIframe = document.getElementById('powerAppIframe'); // Get the iframe by its ID
//         const powerAppId = '3243308d-d91c-4948-a5e3-e98e3a7d8ae5'; // Your specific Power App ID

//         // Only try to embed if we got essential Teams context
//         if (powerAppIframe && teamId !== 'Not available' && channelId !== 'Not available') {
//             let powerAppUrl = `https://apps.powerapps.com/play/${powerAppId}?source=website`;

//             // Append each Teams context parameter to the URL
//             // It's good practice to prefix them or use distinct names to avoid clashes
//             // with Power Apps' internal parameters. E.g., 'teamsTenantId' instead of 'tenantId'.
//             // I'll use simple names for now for clarity, but keep this in mind.
//             powerAppUrl += `&tenantId=${encodeURIComponent(tenantId)}`; // Use tenantId for the Power App to read
//             powerAppUrl += `&teamId=${encodeURIComponent(teamId)}`;
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
    logDebug('Starting Teams SDK initialization...');

    microsoftTeams.app.initialize().then(() => {
        logDebug('Teams SDK initialized successfully');
        return microsoftTeams.app.getContext();
    }).then((context) => {
    if (!context) {
        throw new Error('Teams context is null or undefined');
    }

    logDebug('Teams context received', context);

    // Extract tenant ID with validation
    const tenantId = context.tid || context.user?.tenant?.id || '';
    logDebug('Tenant ID:', tenantId);
        const channelId = context.channel?.id || 'Not available';
        const channelName = context.channel?.displayName || 'Not available';
        const channelType = context.channel?.membershipType || 'Not available'; 

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
    console.log('Hi Jeeva');
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
        // ***** END OF MODIFIED PART *****
    });
});

// Sanitize string for URL (your existing function)
function sanitizeForUrl(str) {
    return str
        .toLowerCase()
        .replace(/[^a-z0-9]+/g, '-')
        .replace(/^-+|-+$/g, '');
}

// Copy URL function (your existing function)
function copyUrl() {
    const urlElement = document.getElementById('sharepointUrl');
    const url = urlElement.textContent;
    
    navigator.clipboard.writeText(url).then(() => {
        const successMsg = document.getElementById('copySuccess');
        successMsg.style.display = 'block';
}).catch(error => {
    logError('Teams initialization or context error', error);
    document.getElementById('sharepointUrl').textContent = 'Error: Failed to initialize Teams or get context';
});

// Sanitize string for URL (your existing function)
function sanitizeForUrl(str) {
    return str
        .toLowerCase()
        .replace(/[^a-z0-9]+/g, '-')
        .replace(/^-+|-+$/g, '');
}

// Copy URL function (your existing function)
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

// Start the app
initializeTeamsApp();

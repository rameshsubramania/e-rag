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


// Initialize Teams SDK
microsoftTeams.app.initialize().then(() => {
    // Get context
    microsoftTeams.app.getContext().then((context) => {
        // Debug information
        console.log('Teams Context:', JSON.stringify(context, null, 2));

        // Extract tenant ID (from context.tid or context.user.tenant.id if available)
        // Note: context.user.userPrincipalName might not always be present or reliable for tenantId directly.
        // For actual tenant ID, context.tid is more direct or context.user.tenant.id
        const tenantId = context.tid || context.user?.tenant?.id || '';
        const tenantName = context.user?.userPrincipalName?.split('@')[1]?.split('.')[0] || ''; // For display, as you had it
        
        document.getElementById('tenantName').textContent = tenantName || 'Not available';
        
        const teamId = context.team?.internalId || 'Not available';
        const teamName = context.team?.displayName || 'Not available';
        const channelId = context.channel?.id || 'Not available';
        const channelName = context.channel?.displayName || 'Not available';
        const channelType = context.channel?.membershipType || 'Not available'; // Use membershipType for accurate private/standard

        // Update UI with values
        document.getElementById('teamId').textContent = teamId;
        document.getElementById('teamName').textContent = teamName;
        document.getElementById('channelId').textContent = channelId;
        document.getElementById('channelName').textContent = channelName;
        document.getElementById('channelType').textContent = channelType; // Display the actual type


        // Generate SharePoint URL (your existing logic)
        if (teamName !== 'Not available' && channelName !== 'Not available') {
            const sanitizedTeamName = sanitizeForUrl(teamName);
            const sanitizedChannelName = sanitizeForUrl(channelName);
            
            console.log('SharePoint site info:', JSON.stringify(context.sharePointSite, null, 2));

            // Use the SharePoint URLs directly from Teams context
            if (context.channel?.membershipType === 'Private') {
                const sharepointUrl = context.sharePointSite.teamSiteUrl + '/Shared%20Documents';
                console.log('Using private channel URL:', sharepointUrl);
                document.getElementById('sharepointUrl').textContent = sharepointUrl;
                // document.getElementById('channelType').textContent = 'Private Channel'; // Already set above
            } else {
                // For standard channels, use the relative URL provided by Teams
                const sharepointUrl = context.sharePointSite.teamSiteUrl + '/Shared%20Documents';
                console.log('Using standard channel URL:', sharepointUrl);
                document.getElementById('sharepointUrl').textContent = sharepointUrl;
                // document.getElementById('channelType').textContent = 'Standard Channel'; // Already set above
            }
        } else {
            document.getElementById('sharepointUrl').textContent = 'Cannot generate URL - missing team or channel name';
        }

        // *** NEW LOGIC TO EMBED POWER APP WITH CONTEXT ***
        const powerAppIframe = document.getElementById('powerAppIframe');
        const powerAppId = '3243308d-d91c-4948-a5e3-e98e3a7d8ae5'; // Your Power App ID

        if (powerAppIframe && teamId !== 'Not available' && channelId !== 'Not available') {
            let powerAppUrl = `https://apps.powerapps.com/play/${powerAppId}?source=website`;

            // Append all context parameters to the Power App URL
            powerAppUrl += `&tenantId=${encodeURIComponent(tenantId)}`; // Using tenantId from context.tid
            powerAppUrl += `&teamId=${encodeURIComponent(teamId)}`;
            powerAppUrl += `&teamName=${encodeURIComponent(teamName)}`;
            powerAppUrl += `&channelId=${encodeURIComponent(channelId)}`;
            powerAppUrl += `&channelName=${encodeURIComponent(channelName)}`;
            powerAppUrl += `&channelType=${encodeURIComponent(channelType)}`; // Pass the actual type

            powerAppIframe.src = powerAppUrl;
            console.log('Power App Embed URL:', powerAppUrl); // For debugging
        } else {
            console.warn('Could not embed Power App: Missing iframe element or essential Teams context.');
        }
    });
});

// Sanitize string for URL
function sanitizeForUrl(str) {
    return str
        .toLowerCase()
        .replace(/[^a-z0-9]+/g, '-') // Replace special chars with hyphens
        .replace(/^-+|-+$/g, ''); // Remove leading/trailing hyphens
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
        console.error('Failed to copy URL:', err);
    });
}
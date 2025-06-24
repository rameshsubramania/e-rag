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
// microsoftTeams.app.initialize().then(() => {
//     // Get context
//     microsoftTeams.app.getContext().then((context) => {
//         // Debug information
//         console.log('Teams Context:', JSON.stringify(context, null, 2));

//         // Extract tenant ID (context.tid is more direct for tenant ID)
//         const tenantId = context.tid || context.user?.tenant?.id || '';
//         // Use userPrincipalName for tenantName display as you had it
//         const tenantNameDisplay = context.user?.userPrincipalName?.split('@')[1]?.split('.')[0] || ''; 
        
//         document.getElementById('tenantName').textContent = tenantNameDisplay || 'Not available';
        
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
//             powerAppUrl += `&teamName=${encodeURIComponent(teamName)}`;
//             powerAppUrl += `&channelId=${encodeURIComponent(channelId)}`;
//             powerAppUrl += `&channelName=${encodeURIComponent(channelName)}`;
//             powerAppUrl += `&channelType=${encodeURIComponent(channelType)}`;

//             // Set the iframe's src, which will load your Power App with these parameters
//             powerAppIframe.src = powerAppUrl;
//             console.log('Power App Embed URL set:', powerAppUrl); // Log for debugging
//         } else {
//             console.warn('Could not embed Power App: Missing iframe element or essential Teams context. Power App will not load with parameters.');
//         }
//         // ***** END OF NEW / MODIFIED PART *****
//     });
// });

// // Sanitize string for URL (your existing function)
// function sanitizeForUrl(str) {
//     return str
//         .toLowerCase()
//         .replace(/[^a-z0-9]+/g, '-') // Replace special chars with hyphens
//         .replace(/^-+|-+$/g, ''); // Remove leading/trailing hyphens
// }

// // Copy URL function (your existing function)
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

        const tenantId = context.tid || context.user?.tenant?.id || '';
        const tenantNameDisplay = context.user?.userPrincipalName?.split('@')[1]?.split('.')[0] || ''; 
        
        document.getElementById('tenantName').textContent = tenantNameDisplay || 'Not available';
        
        const teamId = context.team?.internalId ||context.team?.id || 'Not available';
        const teamName = context.team?.displayName || 'Not available';
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
            
            powerAppUrl += `&FirstName=${encodeURIComponent("Jeeva")}`;
            powerAppUrl += `&tenantId=${encodeURIComponent(tenantId)}`;
            powerAppUrl += `&teamId=${encodeURIComponent(teamId)}`;
            powerAppUrl += `&teamName=${encodeURIComponent(teamName)}`;
            powerAppUrl += `&channelId=${encodeURIComponent(channelId)}`;
            powerAppUrl += `&channelName=${encodeURIComponent(channelName)}`;
            powerAppUrl += `&channelType=${encodeURIComponent(channelType)}`;
            
            // --- ADD THE SHAREPOINT URL PARAMETER HERE ---
            console.log('Power App Embed URL set:', powerAppUrl);
            console.log('Hi Jeeva');
            console.log('Debug: Value of generatedSharepointUrl BEFORE append:', generatedSharepointUrl);

    powerAppUrl += `&sharepointUrl=${encodeURIComponent(generatedSharepointUrl)}`; // THIS LINE


            powerAppIframe.src = powerAppUrl;
            console.log('Power App Embed URL set:', powerAppUrl);
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
        setTimeout(() => {
            successMsg.style.display = 'none';
        }, 3000);
    }).catch(err => {
        console.error('Failed to copy URL:', err);
    });
}

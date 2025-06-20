// Initialize Teams SDK
microsoftTeams.app.initialize().then(() => {
    // Get context
    microsoftTeams.app.getContext().then((context) => {
        // Debug information
        console.log('Teams Context:', JSON.stringify(context, null, 2));
        // Extract tenant name from user's principal name
        const tenantName = context.user?.userPrincipalName?.split('@')[1]?.split('.')[0] || '';
        document.getElementById('tenantName').textContent = tenantName || 'Not available';
        const teamId = context.team?.internalId || 'Not available';
        const teamName = context.team?.displayName || 'Not available';
        const channelId = context.channel?.id || 'Not available';
        const channelName = context.channel?.displayName || 'Not available';

        // Update UI with values
        document.getElementById('teamId').textContent = teamId;
        document.getElementById('teamName').textContent = teamName;
        document.getElementById('channelId').textContent = channelId;
        document.getElementById('channelName').textContent = channelName;

        // Generate SharePoint URL
        if (teamName !== 'Not available' && channelName !== 'Not available') {
            const sanitizedTeamName = sanitizeForUrl(teamName);
            const sanitizedChannelName = sanitizeForUrl(channelName);
            
            console.log('SharePoint site info:', JSON.stringify(context.sharePointSite, null, 2));

            // Use the SharePoint URLs directly from Teams context
            if (context.channel?.membershipType === 'Private') {
                const sharepointUrl = context.sharePointSite.teamSiteUrl + '/Shared%20Documents';
                console.log('Using private channel URL:', sharepointUrl);
                document.getElementById('sharepointUrl').textContent = sharepointUrl;
                document.getElementById('channelType').textContent = 'Private Channel';
            } else {
                // For standard channels, use the relative URL provided by Teams
                const sharepointUrl = context.sharePointSite.teamSiteUrl + '/Shared%20Documents';
                console.log('Using standard channel URL:', sharepointUrl);
                document.getElementById('sharepointUrl').textContent = sharepointUrl;
                document.getElementById('channelType').textContent = 'Standard Channel';
            }
        } else {
            document.getElementById('sharepointUrl').textContent = 'Cannot generate URL - missing team or channel name';
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

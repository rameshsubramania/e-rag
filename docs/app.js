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
            
            // Debug channel type
            console.log('Channel type:', context.channel?.type);
            console.log('Channel properties:', JSON.stringify(context.channel, null, 2));

            // Check if it's a private channel - checking both type and membershipType
            if (context.channel?.type === 'Private' || context.channel?.membershipType === 'private') {
                // Private channel URL format
                // For private channels, we'll use the team name and channel name combination
                const privateTeamName = sanitizeForUrl(teamName);
                const privateChannelName = sanitizeForUrl(channelName);
                const sharepointUrl = `https://${tenantName}.sharepoint.com/sites/${privateTeamName}-${privateChannelName}/Shared%20Documents`;
                console.log('Generated private channel URL:', sharepointUrl);
                document.getElementById('sharepointUrl').textContent = sharepointUrl;
                document.getElementById('channelType').textContent = 'Private Channel';
            } else {
                // Regular channel URL format - points to the channel's folder in the team site
                const sharepointUrl = `https://${tenantName}.sharepoint.com/sites/${sanitizedTeamName}/Shared%20Documents/${channelName}`;
                console.log('Generated standard channel URL:', sharepointUrl);
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

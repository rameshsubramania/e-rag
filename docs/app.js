// Initialize Teams SDK
microsoftTeams.app.initialize().then(() => {
    // Get context
    microsoftTeams.app.getContext().then((context) => {
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
            
            // Check if it's a private channel
            if (context.channel?.type === 'Private') {
                // Private channel URL format
                const sharepointUrl = `https://axleinfo.sharepoint.com/sites/${sanitizedTeamName}-${sanitizedChannelName}/Documents`;
                document.getElementById('sharepointUrl').textContent = sharepointUrl;
                document.getElementById('channelType').textContent = 'Private Channel';
            } else {
                // Regular channel URL format - points to the channel's folder in the team site
                const sharepointUrl = `https://axleinfo.sharepoint.com/sites/${sanitizedTeamName}/Shared%20Documents/General/${channelName}`;
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

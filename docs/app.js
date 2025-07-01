


document.addEventListener('DOMContentLoaded', function () {
    // Initialize Microsoft Teams SDK
    microsoftTeams.app
      .initialize()
      .then(() => microsoftTeams.app.getContext())
      .then((context) => {
        console.log('Teams Context:', JSON.stringify(context, null, 2));
  
        const tenantName = context.user?.userPrincipalName?.split('@')[1]?.split('.')[0] || '';
        const teamId = context.team?.internalId || 'Not available';
        const teamName = context.team?.displayName || 'Not available';
        var channelId = context.channel?.id || 'Not available';
        var channelName = context.channel?.displayName || 'Not available';
        var channelType = context.channel?.membershipType || 'Unknown';
  
        // Generate SharePoint URL
        var sharepointUrl = 'Not available';
        if (
          teamName !== 'Not available' &&
          channelName !== 'Not available' &&
          context.sharePointSite?.teamSiteUrl
        ) {
          if (channelType === 'Private') {
            sharepointUrl = `${context.sharePointSite.teamSiteUrl}/Shared%20Documents`;
          } else {
            const encodedChannelName = encodeURIComponent(channelName);
            sharepointUrl = `${context.sharePointSite.teamSiteUrl}/Shared%20Documents/${encodedChannelName}`;
          }
        } else {
          sharepointUrl = 'Cannot generate URL - missing team or channel name';
        }
  
        const sharepointLabel = document.getElementById('sharepointUrl');
        if (sharepointLabel) {
          sharepointLabel.textContent = sharepointUrl;
        }
      })
      .catch((error) => {
        console.error('Error initializing Teams SDK:', error);
      });
      console.log('Outside SharePoint URL',sharepointUrl);
    // Handle agent creation button
    const createAgentBtn = document.getElementById('createAgentBtn');
    createAgentBtn.addEventListener('click', createAgent);
  
    createAgentBtn.addEventListener('mouseenter', function () {
      this.style.transform = 'translateY(-2px)';
      this.style.boxShadow = '0 6px 16px rgba(121, 80, 242, 0.2)';
    });
  
    createAgentBtn.addEventListener('mouseleave', function () {
      this.style.transform = 'translateY(0)';
      this.style.boxShadow = 'none';
    });
  
    const loginBtn = document.getElementById('loginBtn');
    if (loginBtn) {
      loginBtn.style.display = 'none';
    }
  
    showNotification('✅ App loaded successfully!');
  });
  
  // Function to show notification
  function showNotification(message, isError = false) {
    let notification = document.getElementById('notification');
    if (!notification) {
      notification = document.createElement('div');
      notification.id = 'notification';
      document.body.appendChild(notification);
    }
  
    notification.textContent = message;
    notification.style.backgroundColor = isError ? '#f44336' : '#4CAF50';
    notification.style.transform = 'translateX(0)';
  
    setTimeout(() => {
      notification.style.transform = 'translateX(120%)';
    }, 5000);
  }
  
  // Function to create agent by calling the Logic App
  async function createAgent() {
    const agentName = document.getElementById('agentName').value.trim();
    const model = document.getElementById('modelSelect').value;
   
  console.log('Inside SharePoint URL',sharepointUrl);

    if (!agentName) {
      showNotification('Please enter a name for your agent', true);
      return;
    }
  
    const createAgentBtn = document.getElementById('createAgentBtn');
    const originalText = createAgentBtn.textContent;
  
    try {
      createAgentBtn.disabled = true;
      createAgentBtn.textContent = 'Creating...';
  
      showNotification(`Creating agent "${agentName}" with model ${model}...`);
  
      const url = 'https://prod-41.westus.logic.azure.com:443/workflows/e5f0ce23f3ea415696da0d9b4eeed2ec/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=IZXxoQiXyN8FToQ0GSaFPAy8iO9NEDf9vx5qRP7g0NA';
  
      const requestBody = {
        botName: agentName,
        botModel: model,
        url: sharepointUrl,
        cname: channelName,
        cid: channelId,
        timestamp: new Date().toISOString(),
      };
  
      const response = await fetch(url, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(requestBody),
      });
  
      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }
  
      const data = await response.json();
      console.log('Flow response:', data);
      console.log('Flow response:', sharepointUrl);

  
      showNotification(`✅ Agent "${agentName}" created successfully!`);
      document.getElementById('agentName').value = '';
    } catch (error) {
      console.error('Error creating agent:', error);
      showNotification(`❌ Failed to create agent: ${error.message}`, true);
    } finally {
      createAgentBtn.disabled = false;
      createAgentBtn.textContent = originalText;
    }
  }
  
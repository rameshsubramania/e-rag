// Declare globals at the very top of the script so they exist everywhere
let sharepointUrlBuild = '';
let channelName = '';
let channelId = '';

document.addEventListener('DOMContentLoaded', function () {
  // Initialize Microsoft Teams SDK
  microsoftTeams.app
    .initialize()
    .then(() => microsoftTeams.app.getContext())
    .then((context) => {
      console.log('Teams Context:', JSON.stringify(context, null, 2));

      const teamName = context.team?.displayName || 'Not available';
      channelId = context.channel?.id || 'Not available';
      channelName = context.channel?.displayName || 'Not available';
      const channelType = context.channel?.membershipType || 'Unknown';

      // Build SharePoint URL
      if (
        teamName !== 'Not available' &&
        channelName !== 'Not available' &&
        context.sharePointSite?.teamSiteUrl
      ) {
        if (channelType === 'Private') {
          sharepointUrlBuild = `${context.sharePointSite.teamSiteUrl}/Shared%20Documents`;
        } else {
          const encodedChannelName = encodeURIComponent(channelName);
          sharepointUrlBuild = `${context.sharePointSite.teamSiteUrl}/Shared%20Documents/${encodedChannelName}`;
        }
      } else {
        sharepointUrlBuild = '';
        console.warn('Cannot generate URL - missing team or channel name.');
      }

      console.log('Initialized SharePoint URL:', sharepointUrlBuild);

      // Optional: Show the URL somewhere if you uncomment the label in HTML
      // const sharepointLabel = document.getElementById('sharepointUrl');
      // if (sharepointLabel) {
      //   sharepointLabel.textContent = sharepointUrlBuild || 'N/A';
      // }

      showNotification('✅ App initialized successfully!');
    })
    .catch((error) => {
      console.error('Error initializing Teams SDK:', error);
      showNotification(`❌ Failed to initialize Teams SDK: ${error.message}`, true);
    });

  // Button logic
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

  // If you have a login button you want to hide:
  const loginBtn = document.getElementById('loginBtn');
  if (loginBtn) {
    loginBtn.style.display = 'none';
  }
});

// Function to show success screen
function showSuccessScreen(agentName, model) {
  // Hide initial screen
  document.getElementById('initialScreen').style.display = 'none';
  
  // Show success screen
  const successScreen = document.getElementById('successScreen');
  successScreen.style.display = 'block';
  
  // Set agent details
  document.getElementById('successAgentName').textContent = agentName;
  document.getElementById('successModel').textContent = model === 'gpt-4' ? 'GPT-4' : 'GPT-3.5 Turbo';
  
  // Add click handler for back button
  document.getElementById('backToCreateBtn').addEventListener('click', function() {
    // Show initial screen
    document.getElementById('initialScreen').style.display = 'block';
    // Hide success screen
    successScreen.style.display = 'none';
    // Reset form
    document.getElementById('agentName').value = '';
  });
}

// Function to create agent
async function createAgent() {
  const agentName = document.getElementById('agentName').value.trim();
  const model = document.getElementById('modelSelect').value;

  if (!agentName) {
    showNotification('Please enter a name for your agent', true);
    return;
  }

  if (!sharepointUrlBuild) {
    showNotification('Cannot create agent: SharePoint URL is not available', true);
    return;
  }

  const createAgentBtn = document.getElementById('createAgentBtn');
  const originalText = createAgentBtn.textContent;

  try {
    createAgentBtn.disabled = true;
    createAgentBtn.textContent = 'Creating...';

    showNotification(`Creating agent "${agentName}" with model ${model}...`);

    const url =
      'https://prod-41.westus.logic.azure.com:443/workflows/e5f0ce23f3ea415696da0d9b4eeed2ec/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=IZXxoQiXyN8FToQ0GSaFPAy8iO9NEDf9vx5qRP7g0NA';

    const requestBody = {
      botName: agentName,
      botModel: model,
      url: sharepointUrlBuild,
      cname: channelName,
      cid: channelId,
      timestamp: new Date().toISOString(),
    };

    console.log('Sending request with:', requestBody);

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

// This parses the JSON into a JS objectAdd commentMore actions
const data = await response.json();
console.log('Flow response:', data);

// Extract the properties
const createdBotName = data.botName;
const createdBotModel = data.botModel;

console.log('Created Bot Name:', createdBotName);
console.log('Created Bot Model:', createdBotModel);

// You can also show it in a notification
showNotification(`✅ Agent "${createdBotName}" (${createdBotModel}) created successfully!`);
    
  } catch (error) {
    console.error('Error creating agent:', error);
    showNotification(`❌ Failed to create agent: ${error.message}`, true);
  } finally {
    createAgentBtn.disabled = false;
    createAgentBtn.textContent = originalText;
  }
}

// Function to show notifications
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

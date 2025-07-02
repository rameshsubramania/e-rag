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

//New Status Code

// Your Logic App flow URL
const flowUrl = 'https://prod-66.westus.logic.azure.com:443/workflows/ae73ec5a5772423cb733a1860271241c/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=MC48I55t5lRY9EewVtiHSxwcDsRwUGVArQbWrVZjYGU'; // Replace with your actual URL

// Handle Create Agent button click
document.getElementById('createAgentBtn').addEventListener('click', async function () {
  // Get form values
  const agentName = document.getElementById('agentName').value.trim();
  const model = document.getElementById('modelSelect').value;
  const sharepointUrl = document.getElementById('sharepointUrl').value.trim();
  const channelName = document.getElementById('channelName').value.trim();
  const channelId = document.getElementById('channelId').value.trim();

  if (!agentName) {
    alert('Please enter an agent name.');
    return;
  }

  // Disable the button to prevent multiple clicks
  this.disabled = true;

  // Immediately show waiting screen
  showWaitingScreen(agentName, model);

  // Start polling
  pollStatusUntilSuccess(agentName, model, sharepointUrl, channelName, channelId);
});

// Function to show the "waiting" screen
function showWaitingScreen(agentName, model) {
  document.getElementById('initialScreen').style.display = 'none';
  const successScreen = document.getElementById('successScreen');
  successScreen.style.display = 'block';

  // Initially show "Creating your agent..."
  successScreen.querySelector('h2').textContent = 'Creating your agent...';
  successScreen.querySelector('p').textContent = 'Please wait while we set things up.';
  
  document.getElementById('successAgentName').textContent = agentName;
  document.getElementById('successModel').textContent = model === 'gpt-4' ? 'GPT-4' : 'GPT-3.5 Turbo';
}

// Function to show the final success message
function showSuccessScreen(agentName, model) {
  const successScreen = document.getElementById('successScreen');
  successScreen.querySelector('h2').textContent = 'Agent Created Successfully!';
  successScreen.querySelector('p').textContent = 'Your AI agent is now ready to use. You can access it from your Teams chat.';

  document.getElementById('successAgentName').textContent = agentName;
  document.getElementById('successModel').textContent = model === 'gpt-4' ? 'GPT-4' : 'GPT-3.5 Turbo';
}

// Function to poll until success
async function pollStatusUntilSuccess(agentName, model, sharepointUrl, channelName, channelId) {
  const url = "https://prod-66.westus.logic.azure.com:443/workflows/ae73ec5a5772423cb733a1860271241c/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=MC48I55t5lRY9EewVtiHSxwcDsRwUGVArQbWrVZjYGU";

  const requestBody = {
    botName: agentName,
    botModel: model,
    url: sharepointUrl,
    cname: channelName,
    cid: channelId,
    timestamp: new Date().toISOString(),
  };

  let keepPolling = true;
  const statusElement = document.getElementById('successScreen').querySelector('p');

  while (keepPolling) {
    console.log("⏳ Checking status...");
    statusElement.textContent = 'Checking status...';

    try {
      const response = await fetch(url, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(requestBody),
      });

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      const data = await response.json();
      console.log("✅ Received response:", data);

      if (data.Status === "Success") {
        console.log("🎉 Success! Agent is ready.");
        statusElement.textContent = 'Agent is ready to use!';
        showSuccessScreen(agentName, model);
        keepPolling = false;
      } else {
        console.log("🕒 Not ready yet, checking again in 15 seconds...");
        statusElement.textContent = 'Setting up your agent. This may take a few minutes...';
        await new Promise((resolve) => setTimeout(resolve, 15000));
      }
    } catch (error) {
      console.error("❌ Error during polling:", error.message);
      statusElement.textContent = 'Temporary connection issue, retrying...';
      await new Promise((resolve) => setTimeout(resolve, 15000));
    }
  }
}

// Function to create agent
async function createAgent() {
  const agentName = document.getElementById('agentName').value.trim();
  const model = document.getElementById('modelSelect').value;
  const createAgentBtn = document.getElementById('createAgentBtn');
  const originalText = createAgentBtn.textContent;

  if (!agentName) {
    showNotification('Please enter a name for your agent', true);
    return;
  }

  if (!sharepointUrlBuild) {
    showNotification('Cannot create agent: SharePoint URL is not available', true);
    return;
  }

  try {
    // Show waiting screen immediately
    showWaitingScreen(agentName, model);
    createAgentBtn.disabled = true;
    createAgentBtn.textContent = 'Creating...';

    const url = 'https://prod-41.westus.logic.azure.com:443/workflows/e5f0ce23f3ea415696da0d9b4eeed2ec/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=IZXxoQiXyN8FToQ0GSaFPAy8iO9NEDf9vx5qRP7g0NA';

    const requestBody = {
      botName: agentName,
      botModel: model,
      url: sharepointUrlBuild,
      cname: channelName,
      cid: channelId,
      timestamp: new Date().toISOString(),
    };

    console.log('Sending request with:', requestBody);

    // Make the API call
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

    // If you want to show the success screen after API call completes
    showSuccessScreen(agentName, model);
    
    // If you want to start polling for status instead of showing success immediately
    // pollStatusUntilSuccess(agentName, model, sharepointUrlBuild, channelName, channelId);

  } catch (error) {
    console.error('Error creating agent:', error);
    showNotification(`❌ Failed to create agent: ${error.message}`, true);
    // Hide waiting screen and show the form again on error
    document.getElementById('initialScreen').style.display = 'block';
    document.getElementById('successScreen').style.display = 'none';
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

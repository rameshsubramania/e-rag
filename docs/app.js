// Declare globals at the very top of the script so they exist everywhere
let sharepointUrlBuild = '';
let channelName = '';
let channelId = '';
let currentBotName = '';
let currentBotModel = '';

// Chat context variables
let currentAgentName = '';
let currentModel = '';
let currentSharepointUrl = '';
let currentChannelName = '';
let currentChannelId = '';

// Function to check bot existence and route accordingly
async function checkBotExistence() {
  try {
    // Prepare the request body with all required parameters
    const requestBody = {
      botName: currentAgentName,
      botModel: currentModel,
      url: sharepointUrlBuild,
      cname: channelName,
      cid: channelId,
      timestamp: new Date().toISOString()
    }; 

    console.log('Sending bot existence check with:', requestBody);
    addLog('info', 'Checking bot existence', requestBody);
    
    // Make API call to check if bot exists for this channel
    const response = await fetch('https://prod-143.westus.logic.azure.com:443/workflows/c10edf5d105a4506b13cd787bb50b1b4/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=s4eBbE9niGQBJq_QK_rmyk-ASgEE3Q-8RF3fVUtXfnk', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(requestBody),
    });

    if (!response.ok) {
      throw new Error(`API request failed with status ${response.status}`);
    }

    const data = await response.json();
    console.log('Bot existence check response:', data);
    addLog('info', 'Bot existence check response', data);
    
    if (data.bot === 'Exist') {
      // Bot exists, show screen5 (chat screen) with existing bot
      currentBotName = data.botName || currentAgentName;
      console.log('Bot exists, showing screen5 with:', {
        botName: currentBotName,
        model: currentModel,
        url: sharepointUrlBuild,
        channelName: channelName,
        channelId: channelId
      });
      
      // Show screen5 (chat screen)
      showScreen5(
        currentBotName, 
        currentModel, 
        sharepointUrlBuild, 
        channelName, 
        channelId
      );
      
      return true;
    } else if (data.bot === 'Not Exist') {
      // Bot doesn't exist, show firstScreen (bot creation)
      showFirstScreen();
      return false;
    } else {
      // Unexpected response
      throw new Error('Unexpected response from bot existence check');
    }
  } catch (error) {
    console.error('Error getting bot response:', error);
    addLog('error', 'Chat API request failed', { error: error.message });
    // If there's an error, show firstScreen as a fallback
    showFirstScreen();
    return false;
  }
}

// Function to show screen5 (chat screen)
function showScreen5(botName, botModel, sharepointUrl, channelName, channelId) {
  console.log('showScreen5 called with:', { botName, botModel, sharepointUrl, channelName, channelId });
  
  try {
    // Hide all other screens
    hideAllScreens();
    
    // Show screen5
    const screen5 = document.getElementById('screen5');
    if (!screen5) {
      throw new Error('Screen5 element not found');
    }
    
    screen5.style.display = 'flex';
    
    // Update UI elements in screen5
    const headerTitle = screen5.querySelector('.header-title');
    if (headerTitle) {
      headerTitle.textContent = botName || 'Chat Assistant';
    }
    
    // Store values for later use
    currentBotName = botName;
    currentBotModel = botModel;
    sharepointUrlBuild = sharepointUrl;
    currentChannelName = channelName;
    currentChannelId = channelId;
    
    // Initialize chat functionality for screen5
    initializeScreen5Chat(botName, botModel);
    
    console.log('Screen5 (chat) is now visible');
  } catch (error) {
    console.error('Error in showScreen5:', error);
    showNotification('Error initializing chat. Please refresh the page.', true);
    showFirstScreen();
  }
}

// Function to show firstScreen (creation screen)
function showFirstScreen() {
  hideAllScreens();
  const container = document.querySelector('.container');
  if (container) {
    container.style.display = 'flex';
  }
}

// Function to show secondScreen
function showSecondScreen() {
  hideAllScreens();
  const secondScreen = document.getElementById('secondScreen');
  if (secondScreen) {
    secondScreen.classList.add('active');
  }
}

// Function to show thirdScreen
function showThirdScreen() {
  hideAllScreens();
  const thirdScreen = document.getElementById('thirdScreen');
  if (thirdScreen) {
    thirdScreen.classList.add('active');
  }
}

// Function to show fourthScreen
function showFourthScreen() {
  hideAllScreens();
  const fourthScreen = document.getElementById('fourthScreen');
  if (fourthScreen) {
    fourthScreen.classList.add('active');
    fourthScreen.style.display = 'flex';
  }
  // Start the step animation
  animateProcessingSteps();
}

// Function to hide all screens
function hideAllScreens() {
  // Hide container screens
  const container = document.querySelector('.container');
  if (container) {
    container.style.display = 'none';
  }
  
  // Hide all screen elements
  const screens = ['secondScreen', 'thirdScreen', 'fourthScreen', 'screen5'];
  screens.forEach(screenId => {
    const screen = document.getElementById(screenId);
    if (screen) {
      screen.style.display = 'none';
      screen.classList.remove('active');
    }
  });
}

// Initialize the application
async function initializeApp() {
  try {
    // Initialize Microsoft Teams SDK
    await microsoftTeams.app.initialize();
    
    // Show loading screen
    document.getElementById('loadingScreen').style.display = 'flex';
    document.querySelector('.container').style.display = 'none';
    
    // Get Teams context
    const context = await microsoftTeams.app.getContext();
    
    // Store channel info
    channelName = context.channel?.displayName || '';
    channelId = context.channel?.id || '';
    
    // Log context for debugging
    console.log('Teams Context:', JSON.stringify(context, null, 2));
    
    // Set default values for bot name and model
    currentAgentName = 'Chat Assistant'; // Default name if not provided
    currentModel = 'gpt-4'; // Default model
    
    // Check if bot exists for this channel
    const botExists = await checkBotExistence();
    
    // If bot doesn't exist, we'll show the creation screen
    if (!botExists) {
      // Initialize the rest of the app for bot creation
      initializeBotCreation(context);
    }
  } catch (error) {
    console.error('Error initializing app:', error);
    addLog('error', 'App initialization failed', { error: error.message });
    showCreationScreen();
  }
}

// Logging System
const logs = [];
const maxLogs = 100;

function addLog(level, message, data = null) {
  const timestamp = new Date().toLocaleTimeString();
  const logEntry = {
    timestamp,
    level,
    message,
    data
  };
  
  logs.unshift(logEntry);
  if (logs.length > maxLogs) {
    logs.pop();
  }
  
  updateLogsDisplay();
  
  // Also log to console
  const consoleMethod = level === 'error' ? 'error' : level === 'warning' ? 'warn' : 'log';
  console[consoleMethod](`[${timestamp}] ${message}`, data || '');
}

function updateLogsDisplay() {
  const logsContent = document.getElementById('logsContent');
  if (!logsContent) return;
  
  logsContent.innerHTML = logs.map(log => {
    const dataStr = log.data ? ` - ${JSON.stringify(log.data, null, 2)}` : '';
    return `
      <div class="log-entry">
        <span class="log-timestamp">${log.timestamp}</span>
        <span class="log-level-${log.level}">[${log.level.toUpperCase()}]</span>
        <span class="log-message">${log.message}${dataStr}</span>
      </div>
    `;
  }).join('');
  
  // Auto-scroll to top (newest logs)
  logsContent.scrollTop = 0;
}

function toggleLogsPanel() {
  const logsPanel = document.getElementById('logsPanel');
  const isVisible = logsPanel.style.display === 'flex';
  logsPanel.style.display = isVisible ? 'none' : 'flex';
  
  if (!isVisible) {
    updateLogsDisplay();
  }
}

function clearLogs() {
  logs.length = 0;
  updateLogsDisplay();
  addLog('info', 'Logs cleared');
}

// Initialize the bot creation flow
function initializeBotCreation(context) {
  try {
    // Show the initial screen for bot creation
    showCreationScreen();
    
    // Initialize SharePoint URL builder if needed
    if (context) {
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
      showNotification('✅ App initialized successfully!');
    }
    
    // Set up create agent button event listeners
    const createAgentBtn = document.getElementById('createAgentBtn');
    if (createAgentBtn) {
      createAgentBtn.addEventListener('click', createAgent);
      
      createAgentBtn.addEventListener('mouseenter', function () {
        this.style.transform = 'translateY(-2px)';
        this.style.boxShadow = '0 6px 16px rgba(121, 80, 242, 0.2)';
      });

      createAgentBtn.addEventListener('mouseleave', function () {
        this.style.transform = 'translateY(0)';
        this.style.boxShadow = 'none';
      });
    }

    // Hide login button if it exists
    const loginBtn = document.getElementById('loginBtn');
    if (loginBtn) {
      loginBtn.style.display = 'none';
    }
  } catch (error) {
    console.error('Error initializing bot creation:', error);
    showNotification('Error initializing application. Please refresh and try again.', true);
  }
}

// Function to poll until success
async function pollStatusUntilSuccess(agentName, model, sharepointUrl, channelName, channelId) {
  const url = "https://prod-59.westus.logic.azure.com:443/workflows/09613ec521cb4a438cb7e7df3a1fb99b/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=phnNABFUUeaM5S1hEjhPyMcJaRGR5H8EHPbB11DP_P0";
  const maxAttempts = 2000; // Maximum number of attempts
  let attempt = 1;
  let isSuccess = false;

  const requestBody = {
    botName: agentName,
    botModel: model,
    url: sharepointUrl,
    cname: channelName,
    cid: channelId,
    timestamp: new Date().toISOString(),
  };

  const statusElement = document.getElementById('successScreen').querySelector('p');
  
  // Function to delay between attempts
  const delay = (ms) => new Promise(resolve => setTimeout(resolve, ms));

  while (attempt <= maxAttempts && !isSuccess) {
    try {
      console.log(`⏳ Attempt ${attempt}/${maxAttempts}: Checking agent status...`);
      addLog('debug', `Polling attempt ${attempt}/${maxAttempts}`);
      statusElement.textContent = `Checking agent status (${attempt}/${maxAttempts})...`;

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
      const botResponse = data.botresponse || "I'm sorry, I couldn't process your request at the moment.";
      addLog('info', 'Chat API response received', { userMessage: message, botResponse });
      addLog('debug', `Polling response ${attempt}`, data);

      if (data.Status === "Success" || data.status === "Success") {
        console.log("🎉 Agent is ready!");
        addLog('info', 'Bot creation completed successfully!');
        statusElement.textContent = 'Agent is ready to use!';
        showSuccessScreen(agentName, model, sharepointUrl, channelName, channelId);
        isSuccess = true;
        return; // Exit the function on success
      } else {
        console.log(`Attempt ${attempt}: Agent not ready yet`);
        addLog('debug', `Bot not ready yet (attempt ${attempt})`);
        statusElement.textContent = `Agent is being set up... (${attempt}/${maxAttempts} attempts)`;
      }
    } catch (error) {
      console.error(`❌ Attempt ${attempt} failed:`, error.message);
      addLog('warning', `Polling attempt ${attempt} failed`, { error: error.message });
      statusElement.textContent = `Connection issue, retrying... (${attempt}/${maxAttempts} attempts)`;
    }

    // Only wait if we're going to make another attempt
    if (attempt < maxAttempts) {
      await delay(10000); // Wait 10 seconds before next attempt
    }
    attempt++;
  }

  if (!isSuccess) {
    console.error("❌ Max attempts reached without success");
    addLog('error', 'Bot creation timeout - max attempts reached');
    showNotification('Agent setup is taking longer than expected. Please check back later.', true);
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

    // First API call to create the agent
    const createUrl = 'https://prod-59.westus.logic.azure.com:443/workflows/09613ec521cb4a438cb7e7df3a1fb99b/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=phnNABFUUeaM5S1hEjhPyMcJaRGR5H8EHPbB11DP_P0';
    
    const requestBody = {
      botName: agentName,
      botModel: model,
      url: sharepointUrlBuild,
      cname: channelName,
      cid: channelId,
      timestamp: new Date().toISOString(),
    };

    console.log('Sending create agent request:', requestBody);
    addLog('info', 'Starting bot creation', requestBody);
    console.log('URL:', createUrl);
    
    const response = await fetch(createUrl, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(requestBody),
    });

    if (!response.ok) {
      throw new Error(`Failed to create agent: ${response.status} ${response.statusText}`);
    }

    const responseData = await response.json();
    console.log('Agent creation response:', responseData);
    addLog('info', 'Bot creation API response', responseData);
    
    // Start polling for status after successful creation
    pollStatusUntilSuccess(agentName, model, sharepointUrlBuild, channelName, channelId);
    
  } catch (error) {
    console.error('Error in createAgent:', error);
    addLog('error', 'Bot creation failed', { error: error.message });
    showNotification(`❌ Error: ${error.message}`, true);
    // Show the form again on error
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



// Initialize the app when the DOM is fully loaded
function init() {
  console.log('DOM fully loaded, initializing app...');
  addLog('info', 'Application initializing...');
  
  // Hide all screens initially
  hideAllScreens();
  
  // Initialize the application
  initializeApp().catch(error => {
    console.error('Failed to initialize application:', error);
    addLog('error', 'Critical initialization failure', { error: error.message });
    showNotification('Failed to initialize application. Please refresh the page.', true);
    showFirstScreen();
  });
}

// Check if the DOM is already loaded
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', init);
} else {
  // DOM is already loaded, run immediately
  setTimeout(init, 0);
}

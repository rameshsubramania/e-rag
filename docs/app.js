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
    
    if (data.bot === 'Exist') {
      // Bot exists, show chat screen with existing bot
      currentBotName = data.botName || currentAgentName;
      console.log('Bot exists, showing chat screen with:', {
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
      
      // Debug: Check if chat screen elements exist
      console.log('Chat screen element:', document.getElementById('chatScreen'));
      console.log('Chat agent name element:', document.getElementById('chatAgentName'));
      
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
    console.error('Error checking bot existence:', error);
    showNotification('Error checking bot status. Please try again.', true);
    // If there's an error, show firstScreen as fallback
    showFirstScreen();
    return false;
  }
}

// Function to show the chat screen
function showChatScreen(botName, botModel, sharepointUrl, channelName, channelId) {
  console.log('showChatScreen called with:', { botName, botModel, sharepointUrl, channelName, channelId });
  
  try {
    // Ensure body takes full height
    document.body.style.height = '100%';
    document.documentElement.style.height = '100%';
    
    // Hide loading and initial screens
    const loadingScreen = document.getElementById('loadingScreen');
    const initialScreen = document.getElementById('initialScreen');
    const chatScreen = document.getElementById('chatScreen');
    
    if (!chatScreen) {
      throw new Error('Chat screen element not found');
    }
    
    // Hide other screens
    if (loadingScreen) loadingScreen.style.display = 'none';
    if (initialScreen) initialScreen.style.display = 'none';
    
    // Make sure container is visible and takes full height
    const container = document.querySelector('.container');
    if (container) {
      container.style.display = 'flex';
      container.style.flexDirection = 'column';
      container.style.width = '100%';
      container.style.height = '100%';
      container.style.overflow = 'hidden';
    }
    
    // Show chat screen with proper styling
    chatScreen.style.display = 'flex';
    chatScreen.style.flex = '1';
    chatScreen.style.width = '100%';
    chatScreen.style.height = '100%';
    chatScreen.style.overflow = 'hidden';
    
    // Update UI elements
    const chatAgentNameElement = document.getElementById('chatAgentName');
    const chatAgentNameElement2 = document.getElementById('chatAgentName2');
    const chatModelBadgeElement = document.getElementById('chatModelBadge');
    
    if (!chatAgentNameElement || !chatModelBadgeElement) {
      console.error('Required chat screen elements not found');
      if (initialScreen) initialScreen.style.display = 'block';
      return;
    }
    
    // Set bot info in both header places
    const displayName = botName || 'Chat Assistant';
    chatAgentNameElement.textContent = displayName;
    if (chatAgentNameElement2) {
      chatAgentNameElement2.textContent = displayName;
    }
    chatModelBadgeElement.textContent = botModel === 'gpt-4' ? 'GPT-4' : 'GPT-3.5 Turbo';
    
    // Store values for later use
    currentBotName = botName;
    currentBotModel = botModel;
    sharepointUrlBuild = sharepointUrl;
    
    // Force a reflow to ensure styles are applied
    setTimeout(() => {
      // Initialize chat if the function exists
      if (typeof initializeChat === 'function') {
        try {
          initializeChat(botName, botModel);
          console.log('Chat initialized successfully');
        } catch (error) {
          console.error('Error initializing chat:', error);
        }
      }
      
      // Scroll to bottom of chat
      const chatMessages = document.getElementById('chatMessages');
      if (chatMessages) {
        chatMessages.scrollTop = chatMessages.scrollHeight;
      }
      
      console.log('Chat screen should now be visible');
      console.log('Chat screen dimensions:', {
        width: chatScreen.offsetWidth,
        height: chatScreen.offsetHeight,
        display: window.getComputedStyle(chatScreen).display,
        visibility: window.getComputedStyle(chatScreen).visibility
      });
    }, 0);
  } catch (error) {
    console.error('Error in showChatScreen:', error);
    // Fallback to show error to user
    showNotification('Error initializing chat. Please refresh the page.', true);
    // Try to show creation screen as fallback
    showCreationScreen();
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
  const agentNameInput = document.getElementById('agentName');
  const agentName = agentNameInput.value.trim();

  if (agentName === '') {
    alert('Please give your agent a name before proceeding.');
    agentNameInput.focus();
    agentNameInput.style.borderColor = 'red';
    return;
  }
  
  // Store the agent name and model for later use
  currentAgentName = agentName;
  currentModel = document.getElementById('modelSelect').value;
  
  agentNameInput.style.borderColor = '#D1D5DB'; // Reset border color
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
  
  // Start bot creation polling
  startBotCreation();
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
    currentModel = botModel;
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
    showCreationScreen();
  }
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

// Navigation functions for the new UI
function skipStep() {
  // Handle skip functionality from secondScreen - trigger bot creation
  console.log('Step skipped from secondScreen');
  showFourthScreen();
}

function nextStep() {
  // Navigate to third screen from secondScreen
  document.getElementById('secondScreen').style.display = 'none';
  document.getElementById('thirdScreen').classList.add('active');
}

function goBackToSecond() {
  // Navigate back to second screen from thirdScreen
  document.getElementById('thirdScreen').classList.remove('active');
  document.getElementById('thirdScreen').style.display = 'none';
  document.getElementById('secondScreen').style.display = 'flex';
  document.getElementById('secondScreen').classList.add('active');
}

function skipThirdStep() {
  // Handle skip functionality from thirdScreen - go back to second screen
  goBackToSecond();
}

function confirmSelection() {
  // Handle confirm from thirdScreen - trigger bot creation
  const selectedMethod = document.querySelector('input[name="searchMethod"]:checked');
  if (selectedMethod) {
    console.log('Selected search method:', selectedMethod.value);
    // Navigate to fourth screen (AI Agent Processing)
    showFourthScreen();
  }
}

// Function to start bot creation and polling
function startBotCreation() {
  console.log('Starting bot creation with:', {
    agentName: currentAgentName,
    model: currentModel
  });
  
  // Start polling for bot creation status
  pollStatusUntilSuccess(currentAgentName, currentModel, '', '', '');
}

// Function to animate processing steps in fourthScreen
function animateProcessingSteps() {
  const steps = ['step1', 'step2', 'step3'];
  let currentStep = 0;

  function activateNextStep() {
    if (currentStep < steps.length) {
      // Remove active class from previous step
      if (currentStep > 0) {
        document.getElementById(steps[currentStep - 1]).classList.remove('active');
        document.getElementById(steps[currentStep - 1]).classList.add('completed');
      }
      
      // Add active class to current step
      document.getElementById(steps[currentStep]).classList.add('active');
      
      currentStep++;
      
      // Continue to next step after 2 seconds
      if (currentStep < steps.length) {
        setTimeout(activateNextStep, 2000);
      } else {
        // Complete the last step after 2 seconds
        setTimeout(() => {
          document.getElementById(steps[currentStep - 1]).classList.remove('active');
          document.getElementById(steps[currentStep - 1]).classList.add('completed');
        }, 2000);
      }
    }
  }

  // Start the animation after a short delay
  setTimeout(activateNextStep, 500);
}


// Function to initialize screen5 chat functionality
function initializeScreen5Chat(botName, botModel) {
  const chatMessages = document.getElementById('chat-messages');
  const promptInput = document.getElementById('prompt-input');
  const sendButton = document.getElementById('send-button');
  
  if (!chatMessages || !promptInput || !sendButton) {
    console.error('Required screen5 elements not found');
    return;
  }
  
  // Clear any existing messages and add welcome message
  chatMessages.innerHTML = `
    <div class="initial-greeting">
      <h3>Hi, I'm ${botName}</h3>
      <p>Good Day! How may I assist you today?</p>
    </div>
  `;
  
  // Function to add a message to the chat
  function addMessage(isUser, message) {
    const messageDiv = document.createElement('div');
    messageDiv.className = `message ${isUser ? 'user-message' : 'bot-message'}`;
    messageDiv.textContent = message;
    
    chatMessages.appendChild(messageDiv);
    chatMessages.scrollTop = chatMessages.scrollHeight;
    
    return messageDiv;
  }
  
  // Function to handle sending a message
  async function sendMessage() {
    const message = promptInput.value.trim();
    if (message === '') return;
    
    // Add user message to chat
    addMessage(true, message);
    
    // Clear input
    promptInput.value = '';
    
    // Show typing indicator
    const typingIndicator = addMessage(false, 'Typing...');
    typingIndicator.classList.add('typing-indicator');
    
    try {
      const url = "https://prod-72.westus.logic.azure.com:443/workflows/726b9d82ac464db1b723c2be1bed19f9/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=OYyyRREMa-xCZa0Dut4kRZNoYPZglb1rNXSUx-yMH_U";
      
      try {
        const requestBody = {
          botName: currentAgentName,
          botModel: currentModel,
          url: currentSharepointUrl,
          cname: currentChannelName,
          cid: currentChannelId,
          userMessage: message,
          timestamp: new Date().toISOString(),
        };

        showDebugMessage(`Attempt ${attempt} of ${maxAttempts} to connect to server...`);

      const response = await fetch(url, {
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
      const botResponse = data.botresponse || "I'm sorry, I couldn't process your request at the moment.";
      
      // Remove typing indicator
      typingIndicator.remove();
      
      // Add bot response
      addMessage(false, botResponse);
      
    } catch (error) {
      console.error('Error getting bot response:', error);
      // Remove typing indicator
      typingIndicator.remove();
      
      // Show error message
      addMessage(false, "I'm having trouble connecting to the server. Please try again later.");
    }
  }
  
  // Event listeners
  sendButton.addEventListener('click', sendMessage);
  
  promptInput.addEventListener('keypress', (e) => {
    if (e.key === 'Enter') {
      sendMessage();
    }
  });
  
  // Focus on input field
  promptInput.focus();
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
      console.log(`✅ Attempt ${attempt}:`, data);

      if (data.bot === 'Exist') {
        // Bot creation successful!
        console.log('Bot creation completed successfully!');
        
        // Update current bot name if provided
        if (data.botName) {
          currentBotName = data.botName;
        }
        
        // Show screen5 (chat screen) directly
        showScreen5(agentName, model, sharepointUrl, channelName, channelId);
        return true;
      } else {
        console.log(`Attempt ${attempt}: Agent not ready yet`);
        statusElement.textContent = `Agent is being set up... (${attempt}/${maxAttempts} attempts)`;
      }
    } catch (error) {
      console.error(`❌ Attempt ${attempt} failed:`, error.message);
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
    statusElement.textContent = 'Agent setup is taking longer than expected. Please check back later.';
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
    
    // Start polling for status after successful creation
    pollStatusUntilSuccess(agentName, model, sharepointUrlBuild, channelName, channelId);
    
  } catch (error) {
    console.error('Error in createAgent:', error);
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
  
  // Hide all screens initially
  hideAllScreens();
  
  // Initialize the application
  initializeApp().catch(error => {
    console.error('Failed to initialize application:', error);
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

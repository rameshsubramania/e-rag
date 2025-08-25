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
      throw new Error(`API call failed with status: ${response.status}`);
    }
    
    const data = await response.json();
    showDebugMessage('Bot existence check response:', false);
    showDebugMessage(JSON.stringify(data, null, 2));
    
    if (data.bot === 'Exist') {
      showDebugMessage('Bot exists, preparing to show chat screen with:', false);
      showDebugMessage(JSON.stringify({
        botName: currentAgentName,
        model: currentModel,
        sharepointUrl: sharepointUrlBuild,
        channelName,
        channelId
      }, null, 2));
      
      // Show chat screen with SharePoint URL
      await showChatScreen(
        data.botName || currentAgentName,
        currentModel,
        sharepointUrlBuild,
        channelName,
        channelId
      );
      
      // Verify screen transition
      const chatScreen = document.getElementById('chatScreen');
      const chatAgent = document.getElementById('chatAgentName');
      
      showDebugMessage('Chat screen elements status:', false);
      showDebugMessage(JSON.stringify({
        chatScreenExists: !!chatScreen,
        chatScreenDisplay: chatScreen?.style.display,
        agentNameExists: !!chatAgent,
        agentNameContent: chatAgent?.textContent
      }, null, 2));
      
      return true;
      
    } else if (data.bot === 'Not Exist') {
      showDebugMessage('Bot does not exist, showing creation screen');
      await showCreationScreen();
      return false;
      
    } else {
      throw new Error(`Unexpected bot status: ${data.bot}`);
    }
    
  } catch (error) {
    let errorMessage = error.message;
    
    if (error.name === 'AbortError') {
      errorMessage = 'Bot existence check timed out after 30 seconds';
    } else if (!navigator.onLine) {
      errorMessage = 'No internet connection available';
    }
    
    showDebugMessage(`Error checking bot existence: ${errorMessage}`, true);
    showNotification('Error checking bot status. Please check debug panel for details.', true);
    
    // If there's an error, show creation screen as a fallback
    await showCreationScreen();
    return false;
  }
}

// Function to show the chat screen
function showChatScreen(botName, botModel, sharepointUrl, channelName, channelId) {
  // Use global SharePoint URL if none provided
  const effectiveSharePointUrl = sharepointUrl || sharepointUrlBuild;
  console.log('showChatScreen called with:', { botName, botModel, sharepointUrl: effectiveSharePointUrl, channelName, channelId });
  showDebugMessage(`Using SharePoint URL: ${effectiveSharePointUrl}`);
  
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

// Function to show the creation screen (first screen)
function showCreationScreen() {
  document.getElementById('loadingScreen').style.display = 'none';
  const container = document.querySelector('.container');
  if (container) {
    container.style.display = 'flex';
    // Hide all other screens
    hideAllScreens();
    // Show first screen (container is the first screen)
    container.style.display = 'flex';
  }
  const chatScreen = document.getElementById('chatScreen');
  if (chatScreen) chatScreen.style.display = 'none';
}

// Helper function to hide all screens
function hideAllScreens() {
  const screens = ['secondScreen', 'thirdScreen', 'fourthScreen', 'chatScreen'];
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
    sharepointUrlBuild = context.sharePointSite?.teamSiteUrl || '';
    
    // Log context for debugging
    console.log('Teams Context:', JSON.stringify(context, null, 2));
    showDebugMessage(`SharePoint URL: ${sharepointUrlBuild}`);
    
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

// Your Logic App flow URL
const flowUrl = 'https://prod-66.westus.logic.azure.com:443/workflows/ae73ec5a5772423cb733a1860271241c/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=MC48I55t5lRY9EewVtiHSxwcDsRwUGVArQbWrVZjYGU';

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

// Function to show the fourth screen (waiting/processing screen)
function showWaitingScreen(agentName, model) {
  // Hide all other screens
  hideAllScreens();
  const container = document.querySelector('.container');
  if (container) container.style.display = 'none';
  
  // Show fourth screen
  const fourthScreen = document.getElementById('fourthScreen');
  if (fourthScreen) {
    fourthScreen.style.display = 'flex';
    fourthScreen.classList.add('active');
    
    // Start the processing animation
    setTimeout(() => {
      if (typeof animateProcessingSteps === 'function') {
        animateProcessingSteps();
      }
    }, 500);
  }
}

// Function to show the final success message - now directly shows chat screen
function showSuccessScreen(agentName, model, sharepointUrl, channelName, channelId) {
  console.log('Agent creation completed, showing chat screen');
  
  // Directly show chat screen instead of success screen
  showChatScreen(agentName, model, sharepointUrl, channelName, channelId);
}

// Function to show the chat screen
function showChatScreen(agentName, model, sharepointUrl, channelName, channelId) {
  console.log('Showing chat screen for agent:', agentName);
  
  // Hide all other screens
  const screens = ['loadingScreen', 'firstScreen', 'secondScreen', 'thirdScreen', 'fourthScreen'];
  screens.forEach(screenId => {
    const screen = document.getElementById(screenId);
    if (screen) {
      screen.style.display = 'none';
      screen.classList.remove('active');
    }
  });
  
  // Hide container screens
  const containers = document.querySelectorAll('.container');
  containers.forEach(container => {
    container.style.display = 'none';
    container.classList.remove('active');
  });
  
  // Show chat screen
  const chatScreen = document.getElementById('chatScreen');
  if (chatScreen) {
    chatScreen.style.display = 'flex';
    chatScreen.classList.add('active');
    
    // Update chat screen with agent info
    const chatAgentName = document.getElementById('chatAgentName');
    const chatModelBadge = document.getElementById('chatModelBadge');
    
    if (chatAgentName) chatAgentName.textContent = agentName || 'AI Assistant';
    if (chatModelBadge) chatModelBadge.textContent = model === 'gpt-4o' ? 'GPT-4o' : 'GPT-3.5 Turbo';
    
    // Initialize chat functionality
    initializeChatFunctionality(agentName, model, sharepointUrl, channelName, channelId);
    
    // Add welcome message
    addChatMessage(`Hello! I'm ${agentName || 'your AI Assistant'}. How can I help you today?`, 'bot');
  }
}

// Function to initialize chat functionality
function initializeChatFunctionality(agentName, model, sharepointUrl, channelName, channelId) {
  const userInput = document.getElementById('userMessageInput');
  const sendButton = document.getElementById('sendMessageBtn');
  
  if (userInput && sendButton) {
    // Remove existing event listeners
    sendButton.replaceWith(sendButton.cloneNode(true));
    const newSendButton = document.getElementById('sendMessageBtn');
    
    // Add click event listener
    newSendButton.addEventListener('click', () => {
      sendChatMessage(agentName, model, sharepointUrl, channelName, channelId);
    });
    
    // Add enter key event listener
    userInput.addEventListener('keypress', (e) => {
      if (e.key === 'Enter') {
        sendChatMessage(agentName, model, sharepointUrl, channelName, channelId);
      }
    });
  }
}

// Function to send chat message
function sendChatMessage(agentName, model, sharepointUrl, channelName, channelId) {
  const userInput = document.getElementById('userMessageInput');
  const message = userInput.value.trim();
  
  if (message === '') return;
  
  // Add user message to chat
  addChatMessage(message, 'user');
  
  // Clear input
  userInput.value = '';
  
  // Show typing indicator
  addTypingIndicator();
  
  // Send message to bot (simulate for now)
  setTimeout(() => {
    removeTypingIndicator();
    addChatMessage('I received your message: "' + message + '". This is a simulated response. The actual bot integration will be implemented when the backend is ready.', 'bot');
  }, 1500);
}

// Function to add message to chat
function addChatMessage(message, sender) {
  const chatMessages = document.getElementById('chatMessages');
  if (!chatMessages) return;
  
  const messageDiv = document.createElement('div');
  messageDiv.className = `message ${sender}-message`;
  messageDiv.textContent = message;
  
  chatMessages.appendChild(messageDiv);
  
  // Scroll to bottom
  chatMessages.scrollTop = chatMessages.scrollHeight;
}

// Function to add typing indicator
function addTypingIndicator() {
  const chatMessages = document.getElementById('chatMessages');
  if (!chatMessages) return;
  
  const typingDiv = document.createElement('div');
  typingDiv.className = 'message bot-message typing-indicator';
  typingDiv.id = 'typingIndicator';
  typingDiv.textContent = 'AI is typing...';
  
  chatMessages.appendChild(typingDiv);
  chatMessages.scrollTop = chatMessages.scrollHeight;
}

// Function to remove typing indicator
function removeTypingIndicator() {
  const typingIndicator = document.getElementById('typingIndicator');
  if (typingIndicator) {
    typingIndicator.remove();
  }
}

// Complete the app integration - all functions are now properly connected

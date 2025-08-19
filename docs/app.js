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
      showDebugMessage('Bot does not exist, showing first screen');
      await showFirstScreen();
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
    
    // If there's an error, show first screen as a fallback
    await showFirstScreen();
    return false;
  }
}

// Function to show the chat screen - now replaced by showFifthScreen
function showChatScreen(agentName, model, sharepointUrl, channelName, channelId) {
  // This function is kept for backward compatibility but now uses the fifth screen
  showFifthScreen(agentName, model, sharepointUrl, channelName, channelId);
}

// Function to show the fifth screen
function showFifthScreen(botName, botModel, sharepointUrl, channelName, channelId) {
  // Use global SharePoint URL if none provided
  const effectiveSharePointUrl = sharepointUrl || sharepointUrlBuild;
  console.log('showFifthScreen called with:', { botName, botModel, sharepointUrl: effectiveSharePointUrl, channelName, channelId });
  showDebugMessage(`Using SharePoint URL: ${effectiveSharePointUrl}`);
  
  try {
    // Ensure body takes full height
    document.body.style.height = '100%';
    document.documentElement.style.height = '100%';
    
    // Hide all screens
    const loadingScreen = document.getElementById('loadingScreen');
    const fifthScreen = document.getElementById('fifthScreen');
    
    // Hide all screens first
    hideAllScreens();
    
    if (!fifthScreen) {
      throw new Error('Fifth screen element not found');
    }
    
    // Hide other screens
    if (loadingScreen) loadingScreen.style.display = 'none';
    
    // Make sure container is visible and takes full height
    const container = document.querySelector('.container');
    if (container) {
      container.style.display = 'flex';
      container.style.flexDirection = 'column';
      container.style.width = '100%';
      container.style.height = '100%';
      container.style.overflow = 'hidden';
    }
    
    // Show fifth screen with proper styling
    fifthScreen.style.display = 'flex';
    fifthScreen.style.flex = '1';
    fifthScreen.style.width = '100%';
    fifthScreen.style.height = '100%';
    fifthScreen.style.overflow = 'hidden';
    
    // Update UI elements
    const chatAgentNameElement = document.getElementById('chatAgentName');
    const chatAgentNameElement2 = document.getElementById('chatAgentName2');
    const chatModelBadgeElement = document.getElementById('chatModelBadge');
    
    if (!chatAgentNameElement || !chatModelBadgeElement) {
      console.error('Required chat screen elements not found');
      showFirstScreen();
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
    // Try to show first screen as fallback
    showFirstScreen();
  }
}

// Function to show the first screen (agent name and model selection)
function showFirstScreen() {
  document.getElementById('loadingScreen').style.display = 'none';
  
  // Hide all screens first
  hideAllScreens();
  
  // Show first screen
  const firstScreen = document.getElementById('firstScreen');
  if (firstScreen) {
    firstScreen.classList.add('active');
  }
  
  // Make sure container is visible
  const container = document.querySelector('.container');
  if (container) {
    container.style.display = 'flex';
  }
}

// Function to hide all screens
function hideAllScreens() {
  // Get all screens
  const screens = document.querySelectorAll('.screen');
  
  // Remove active class from all screens
  screens.forEach(screen => {
    screen.classList.remove('active');
  });
}

// Function to show the second screen
function showSecondScreen() {
  // Get agent name and model from first screen
  const agentName = document.getElementById('agentName').value.trim();
  const model = document.getElementById('modelSelect').value;
  
  // Validate input
  if (!agentName) {
    showNotification('Please enter an agent name.', true);
    return;
  }
  
  // Store values in global variables
  currentAgentName = agentName;
  currentModel = model;
  
  // Hide all screens
  hideAllScreens();
  
  // Show second screen
  const secondScreen = document.getElementById('secondScreen');
  if (secondScreen) {
    secondScreen.classList.add('active');
  }
}

// Function to handle Skip button on second screen
function skipStep() {
  // When Skip is clicked, create the bot and show the fourth screen
  createAgent();
  showFourthScreen();
}

// Function to handle Next button on second screen
function nextStep() {
  // When Next is clicked, proceed to the third screen
  hideAllScreens();
  const thirdScreen = document.getElementById('thirdScreen');
  if (thirdScreen) {
    thirdScreen.classList.add('active');
  }
}

// Function to show the fourth screen (processing screen)
function showFourthScreen() {
  hideAllScreens();
  const fourthScreen = document.getElementById('fourthScreen');
  if (fourthScreen) {
    fourthScreen.classList.add('active');
    // Start the animation for processing steps
    animateProcessingSteps();
  }
}

// Function to animate the processing steps
function animateProcessingSteps() {
  const steps = document.querySelectorAll('.processing-step');
  let currentStep = 0;
  
  // Function to show each step with a delay
  function showNextStep() {
    if (currentStep < steps.length) {
      steps[currentStep].classList.add('active');
      currentStep++;
      setTimeout(showNextStep, 1500); // Show next step after 1.5 seconds
    }
  }
  
  // Start the animation
  showNextStep();
}

// Function to go back to second screen from third screen
function goBackToSecond() {
  hideAllScreens();
  const secondScreen = document.getElementById('secondScreen');
  if (secondScreen) {
    secondScreen.classList.add('active');
  }
}

// Function to skip third screen and create agent
function skipThirdStep() {
  // Skip the third screen, create the bot and show the fourth screen
  createAgent();
  showFourthScreen();
}

// Function to confirm selection on third screen
function confirmSelection() {
  // Get the selected search method
  const searchMethodOptions = document.querySelectorAll('input[name="searchMethod"]');
  let selectedMethod = '';
  
  searchMethodOptions.forEach(option => {
    if (option.checked) {
      selectedMethod = option.value;
    }
  });
  
  // Store the selected method (if needed for future use)
  // You can add this to the agent creation parameters if needed
  
  // Create the agent and proceed to fourth screen
  createAgent();
  showFourthScreen();
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
    
    // If bot doesn't exist, we'll show the first screen
    if (!botExists) {
      // Initialize the rest of the app for bot creation
      initializeBotCreation(context);
      // Show the first screen
      showFirstScreen();
    }
  } catch (error) {
    console.error('Error initializing app:', error);
    showFirstScreen();
    showNotification('Error initializing app. Please refresh the page.', true);
  }
}

// Initialize the bot creation flow
function initializeBotCreation(context = null) {
  try {
    // Initialize SharePoint URL builder if Teams context is available
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
    } else {
      // Set default values when no Teams context is available
      sharepointUrlBuild = '';
      channelId = 'demo-channel-id';
      channelName = 'Demo Channel';
      console.log('Running in standalone mode without Teams context');
    }
    
    // First screen: Agent name and model selection
    const firstNextBtn = document.getElementById('firstNextBtn');
    if (firstNextBtn) {
      firstNextBtn.addEventListener('click', showSecondScreen);
    }
    
    // Second screen: Next and Skip buttons
    const nextBtn = document.getElementById('nextBtn');
    if (nextBtn) {
      nextBtn.addEventListener('click', nextStep);
    }
    
    const skipBtn = document.getElementById('skipBtn');
    if (skipBtn) {
      skipBtn.addEventListener('click', skipStep);
    }
    
    // Third screen: Back, Skip, and Confirm buttons
    const backBtn = document.getElementById('backBtn');
    if (backBtn) {
      backBtn.addEventListener('click', goBackToSecond);
    }
    
    const skipThirdBtn = document.getElementById('skipThirdBtn');
    if (skipThirdBtn) {
      skipThirdBtn.addEventListener('click', skipThirdStep);
    }
    
    const confirmBtn = document.getElementById('confirmBtn');
    if (confirmBtn) {
      confirmBtn.addEventListener('click', confirmSelection);
    }
    
    // Set up model selection
    const modelSelect = document.getElementById('modelSelect');
    if (modelSelect) {
      modelSelect.addEventListener('change', function() {
        currentModel = this.value;
      });
    }
    
    // Set up agent name input
    const agentNameInput = document.getElementById('agentName');
    if (agentNameInput) {
      agentNameInput.addEventListener('input', function() {
        currentAgentName = this.value.trim();
      });
    }
    
    // Set up create agent button event listeners
    const createAgentBtn = document.getElementById('createAgentBtn');
    if (createAgentBtn) {
      createAgentBtn.addEventListener('click', createAgent);
      
      createAgentBtn.addEventListener('mouseenter', function () {
        this.style.transform = 'translateY(-2px)';
        this.style.boxShadow = '0 6px 20px rgba(102, 126, 234, 0.6)';
      });

      createAgentBtn.addEventListener('mouseleave', function () {
        this.style.transform = 'translateY(0)';
        this.style.boxShadow = '0 4px 15px rgba(102, 126, 234, 0.4)';
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

// Function to show the "waiting" screen - now replaced by showFourthScreen
function showWaitingScreen(agentName, model) {
  // This function is kept for backward compatibility but now uses the fourth screen
  showFourthScreen();
}

// Function to show the final success message - now replaced by showFifthScreen
function showSuccessScreen(agentName, model, sharepointUrl, channelName, channelId) {
  // This function is kept for backward compatibility but now uses the fifth screen
  showFifthScreen(agentName, model, sharepointUrl, channelName, channelId);
}

// Function to show the chat screen
function showChatScreen(agentName, model, sharepointUrl, channelName, channelId) {
  // Clear any existing debug logs from screen
  clearDebugLogs();
  
  // Update global context
  currentAgentName = agentName;
  currentModel = model;
  currentSharepointUrl = sharepointUrl || sharepointUrlBuild;
  currentChannelName = channelName || '';
  currentChannelId = channelId || '';

  // Hide loading screen and show container
  document.getElementById('loadingScreen').style.display = 'none';
  document.querySelector('.container').style.display = 'flex';

  // Hide all screens first
  hideAllScreens();
  
  // Show the fifth screen (chat screen)
  const fifthScreen = document.getElementById('fifthScreen');
  if (fifthScreen) {
    fifthScreen.classList.add('active');
  }
  
  // Set the agent name in the UI
  document.querySelectorAll('.chat-header h2, .sidebar-header h3').forEach(el => {
    el.textContent = agentName;
  });
  document.getElementById('chatAgentName2').textContent = agentName;
  
  // Initialize chat functionality
  initializeChat(agentName, model);
}

// Function to initialize chat functionality
function initializeChat(agentName, model) {
  const chatMessages = document.getElementById('chatMessages');
  const userMessageInput = document.getElementById('userMessageInput');
  const sendMessageBtn = document.getElementById('sendMessageBtn');
  
  // Clear any existing messages
  chatMessages.innerHTML = '';
  
  // Add welcome message
  addWelcomeMessage(agentName);
  
  // Function to add welcome message
  function addWelcomeMessage(agentName) {
    const welcomeMessage = `
      <div class="message bot-message welcome-message">
        <div class="message-avatar">
          <div class="avatar">
            <svg width="24" height="24" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
              <path d="M12 2C6.48 2 2 6.48 2 12C2 17.52 6.48 22 12 22C17.52 22 22 17.52 22 12C22 6.48 17.52 2 12 2ZM12 20C7.59 20 4 16.41 4 12C4 7.59 7.59 4 12 4C16.41 4 20 7.59 20 12C20 16.41 16.41 20 12 20Z" fill="currentColor"/>
              <path d="M12 6C9.79 6 8 7.79 8 10C8 12.21 9.79 14 12 14C14.21 14 16 12.21 16 10C16 7.79 14.21 6 12 6ZM12 12C10.9 12 10 11.1 10 10C10 8.9 10.9 8 12 8C13.1 8 14 8.9 14 10C14 11.1 13.1 12 12 12Z" fill="currentColor"/>
              <path d="M12 15C9.33 15 4 16.34 4 19V21H20V19C20 16.34 14.67 15 12 15ZM6 19C6.22 18.28 9.31 17 12 17C14.7 17 17.8 18.29 18 19H6Z" fill="currentColor"/>
            </svg>
          </div>
        </div>
        <div class="message-content">
          <h3>Hi, I'm <span id="chatAgentName2">${agentName}</span></h3>
          <p>Good Day! How may I assist you today?</p>
        </div>
      </div>
    `;
    chatMessages.innerHTML = welcomeMessage;
  }
  
  // Function to add a message to the chat
  function addMessage(isUser, message) {
    const messageDiv = document.createElement('div');
    messageDiv.className = `message ${isUser ? 'user-message' : 'bot-message'}`;
    
    // Create avatar
    const avatarDiv = document.createElement('div');
    avatarDiv.className = 'message-avatar';
    
    const avatar = document.createElement('div');
    avatar.className = 'avatar';
    
    if (isUser) {
      // User avatar (first letter of the name)
      const userInitial = document.createElement('span');
      userInitial.textContent = 'Y';
      avatar.appendChild(userInitial);
    } else {
      // Bot avatar (icon)
      const botIcon = document.createElement('div');
      botIcon.innerHTML = `
        <svg width="20" height="20" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
          <path d="M12 2C6.48 2 2 6.48 2 12C2 17.52 6.48 22 12 22C17.52 22 22 17.52 22 12C22 6.48 17.52 2 12 2ZM12 20C7.59 20 4 16.41 4 12C4 7.59 7.59 4 12 4C16.41 4 20 7.59 20 12C20 16.41 16.41 20 12 20Z" fill="currentColor"/>
          <path d="M12 6C9.79 6 8 7.79 8 10C8 12.21 9.79 14 12 14C14.21 14 16 12.21 16 10C16 7.79 14.21 6 12 6ZM12 12C10.9 12 10 11.1 10 10C10 8.9 10.9 8 12 8C13.1 8 14 8.9 14 10C14 11.1 13.1 12 12 12Z" fill="currentColor"/>
          <path d="M12 15C9.33 15 4 16.34 4 19V21H20V19C20 16.34 14.67 15 12 15ZM6 19C6.22 18.28 9.31 17 12 17C14.7 17 17.8 18.29 18 19H6Z" fill="currentColor"/>
        </svg>
      `;
      avatar.appendChild(botIcon);
    }
    
    avatarDiv.appendChild(avatar);
    
    // Create message content
    const contentDiv = document.createElement('div');
    contentDiv.className = 'message-content';
    
    if (!isUser) {
      const nameElement = document.createElement('h3');
      nameElement.textContent = agentName;
      contentDiv.appendChild(nameElement);
    }
    
    const textElement = document.createElement('p');
    textElement.textContent = message;
    contentDiv.appendChild(textElement);
    
    // Assemble message
    messageDiv.appendChild(avatarDiv);
    messageDiv.appendChild(contentDiv);
    
    // Add to chat
    chatMessages.appendChild(messageDiv);
    
    // Scroll to the bottom of the chat
    chatMessages.scrollTop = chatMessages.scrollHeight;
    
    return messageDiv;
  }
  
  // Function to handle sending a message
  async function sendMessage() {
    const message = userMessageInput.value.trim();
    if (message === '') return;
    
    // Add user message to chat
    addMessage(true, message);
    
    // Clear input
    userMessageInput.value = '';
    
    // Show typing indicator
    const typingIndicator = addMessage(false, '...');
    typingIndicator.id = 'typing-indicator';
    typingIndicator.querySelector('.message-content p').textContent = 'Typing...';
    
    async function tryRequest(attempt = 1, maxAttempts = 3) {
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
        return data.botresponse || "I'm sorry, I couldn't process your request at the moment.";
      } catch (error) {
        if (attempt < maxAttempts) {
          showDebugMessage(`Attempt ${attempt} failed: ${error.message}. Retrying...`);
          await new Promise(resolve => setTimeout(resolve, 2000 * attempt)); // Exponential backoff
          return tryRequest(attempt + 1, maxAttempts);
        }
        throw error;
      }
    }

    try {
      const botResponse = await tryRequest();
      
      // Remove typing indicator
      const indicator = document.getElementById('typing-indicator');
      if (indicator) indicator.remove();
      
      // Add bot response
      addMessage(false, botResponse);
      
    } catch (error) {
      console.error('Error getting bot response:', error);
      showDebugMessage(`Failed to connect to server: ${error.message}`, true);
      
      // Remove typing indicator
      const indicator = document.getElementById('typing-indicator');
      if (indicator) indicator.remove();
      
      // Show error message with more details
      addMessage(false, "I'm having trouble connecting to the server. Please try again later.");
    }
  }
  
  // Event listeners
  sendMessageBtn.addEventListener('click', sendMessage);
  
  userMessageInput.addEventListener('keypress', (e) => {
    if (e.key === 'Enter') {
      sendMessage();
    }
  });
  
  // Quick action button
  document.querySelector('.quick-action-btn').addEventListener('click', () => {
    userMessageInput.value = 'Tell me about the application';
    userMessageInput.focus();
    sendMessage(); // Automatically send the quick action message
  });
  
  // Sidebar actions
  const newChatBtn = document.querySelector('.sidebar-action-btn:first-child');
  const savedPromptsBtn = document.querySelector('.sidebar-action-btn:last-child');
  
  newChatBtn.addEventListener('click', () => {
    // Clear chat messages
    chatMessages.innerHTML = '';
    // Add welcome message
    addWelcomeMessage(agentName);
    // Set active state
    newChatBtn.classList.add('active');
    savedPromptsBtn.classList.remove('active');
  });
  
  savedPromptsBtn.addEventListener('click', () => {
    // In a real app, this would show saved prompts
    alert('Saved prompts feature coming soon!');
    // Set active state
    savedPromptsBtn.classList.add('active');
    newChatBtn.classList.remove('active');
  });
  
  // Set focus to input field
  userMessageInput.focus();
}

// Function to poll until success
async function pollStatusUntilSuccess(botName, botModel, sharepointUrl, channelName, channelId) {
  const maxAttempts = 30; // Max number of polling attempts
  const pollingInterval = 5000; // 5 seconds between polls
  let attempts = 0;
  let isSuccess = false;
  
  // Update the processing step status in the fourth screen
  const processingSteps = document.querySelectorAll('.processing-step');
  if (processingSteps.length >= 2) {
    processingSteps[1].classList.add('active'); // Activate the "Checking status" step
  }

  while (attempts < maxAttempts && !isSuccess) {
    attempts++;
    console.log(`Polling attempt ${attempts}/${maxAttempts}`);
    
    try {
      const statusUrl = 'https://prod-18.westus.logic.azure.com:443/workflows/e8a3f5c0d7a14f6e8d0c7a9c7e4c4f5f/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx';
      
      const requestBody = {
        botName: botName,
        botModel: botModel,
        url: sharepointUrl,
        cname: channelName,
        cid: channelId,
        timestamp: new Date().toISOString(),
      };

      const response = await fetch(statusUrl, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(requestBody),
      });

      if (!response.ok) {
        throw new Error(`Status check failed: ${response.status} ${response.statusText}`);
      }

      const data = await response.json();
      console.log('Status check response:', data);

      if (data.status === 'Success') {
        isSuccess = true;
        
        // Activate the final processing step
        if (processingSteps.length >= 3) {
          processingSteps[2].classList.add('active'); // Activate the "Ready" step
        }
        
        // Wait 2 seconds before showing the fifth screen (chat)
        setTimeout(() => {
          showFifthScreen(botName, botModel, sharepointUrl, channelName, channelId);
        }, 2000);
        
        break;
      } else if (data.status === 'Failed') {
        throw new Error('Agent creation failed. Please try again.');
      } else {
        // Still in progress
        // Wait for the polling interval before next attempt
        await new Promise(resolve => setTimeout(resolve, pollingInterval));
      }
    } catch (error) {
      console.error('Error in polling:', error);
      showNotification(`Error checking status: ${error.message}`, true);
      // Wait before retry
      await new Promise(resolve => setTimeout(resolve, pollingInterval));
    }
  }

  if (!isSuccess) {
    console.error("❌ Max attempts reached without success");
    showNotification('Agent setup is taking longer than expected. Please try again.', true);
    // Go back to first screen if we couldn't create the agent
    showFirstScreen();
  }
}

// Function to show the fifth screen (chat screen)
function showFifthScreen(botName, botModel, sharepointUrl, channelName, channelId) {
  hideAllScreens();
  
  // Show fifth screen
  const fifthScreen = document.getElementById('fifthScreen');
  if (fifthScreen) {
    fifthScreen.classList.add('active');
    
    // Update the agent name and model in the chat header
    const chatAgentNameElement = document.getElementById('chatAgentName');
    const chatModelBadgeElement = document.getElementById('chatModelBadge');
    
    if (chatAgentNameElement) {
      chatAgentNameElement.textContent = botName || 'Chat Assistant';
    }
    
    if (chatModelBadgeElement) {
      chatModelBadgeElement.textContent = botModel === 'gpt-4' ? 'GPT-4' : 'GPT-3.5 Turbo';
    }
    
    // Initialize the chat functionality
    initializeChat(botName, botModel);
    
    // Scroll chat to bottom
    const chatMessages = document.getElementById('chatMessages');
    if (chatMessages) {
      setTimeout(() => {
        chatMessages.scrollTop = chatMessages.scrollHeight;
      }, 0);
    }
  }
}

// Function to create agent
async function createAgent() {
  // Use the global variables that were set when navigating through screens
  const agentName = currentAgentName;
  const model = currentModel;

  if (!agentName) {
    showNotification('Please enter a name for your agent', true);
    return;
  }

  if (!sharepointUrlBuild) {
    showNotification('Cannot create agent: SharePoint URL is not available', true);
    return;
  }

  try {
    // Show processing in the fourth screen
    showFourthScreen();
    
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
    // Show the first screen again on error
    showFirstScreen();
  }
}

// Function to check bot existence
async function checkBotExistence() {
  try {
    const requestBody = {
      botName: currentAgentName,
      botModel: currentModel,
      url: sharepointUrlBuild,
      cname: channelName,
      cid: channelId,
      timestamp: new Date().toISOString()
    };

    const response = await fetch('https://prod-54.westus.logic.azure.com:443/workflows/0e1a0f6f4c1e4e9f8d7c6b5a4e3d2c1b/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(requestBody),
    });

    if (!response.ok) {
      throw new Error(`Failed to check bot existence: ${response.status} ${response.statusText}`);
    }

    const data = await response.json();
    console.log('Bot existence check response:', data);

    if (data.bot === 'Exist') {
      // If bot exists, show the fifth screen (chat screen)
      showFifthScreen(currentAgentName, currentModel, sharepointUrlBuild, channelName, channelId);
      return true;
    } else if (data.bot === 'Not Exist') {
      // If bot doesn't exist, show the first screen
      showFirstScreen();
      return false;
    } else {
      throw new Error(`Unexpected response from bot existence check: ${JSON.stringify(data)}`);
    }
  } catch (error) {
    console.error('Error in checkBotExistence:', error);
    showNotification(`Error checking bot existence: ${error.message}`, true);
    // Show the first screen on error
    showFirstScreen();
    return false;
  }
}

// Function to show notifications with improved visibility
function showNotification(message, isError = false) {
  console.log(`Showing notification: ${message} (isError: ${isError})`);
  
  let notification = document.getElementById('notification');
  if (!notification) {
    notification = document.createElement('div');
    notification.id = 'notification';
    notification.style.position = 'fixed';
    notification.style.top = '20px';
    notification.style.right = '20px';
    notification.style.padding = '15px 20px';
    notification.style.borderRadius = '4px';
    notification.style.color = 'white';
    notification.style.zIndex = '10000';
    notification.style.maxWidth = '80%';
    notification.style.boxShadow = '0 4px 6px rgba(0,0,0,0.1)';
    document.body.appendChild(notification);
  }

  notification.textContent = message;
  notification.style.backgroundColor = isError ? '#f44336' : '#4CAF50';
  notification.style.display = 'block';
  notification.style.transform = 'translateX(0)';

  setTimeout(() => {
    notification.style.transform = 'translateX(120%)';
  }, 5000);
}



// Initialize the app when the DOM is fully loaded
// Function to show debug messages in console only (no visible UI)
function showDebugMessage(message, error = false) {
    const timestamp = new Date().toISOString();
    const logMessage = `[${timestamp}] ${message}`;
    
    // Only log to console - no visible UI elements
    if (error) {
        console.error(logMessage);
        // Only show critical errors as notifications
        showNotification(`Error: ${message}`, true);
    } else {
        console.log(logMessage);
    }
}

// Function to clear any existing debug logs from screen
function clearDebugLogs() {
    // Remove any existing status log overlay
    const statusLog = document.getElementById('statusLog');
    if (statusLog) {
        statusLog.remove();
    }
    
    // Clear debug panel if it exists
    const debugPanel = document.getElementById('debug');
    if (debugPanel) {
        debugPanel.innerHTML = '';
    }
}

// Function to manually clear chat and debug logs
function clearChatAndLogs() {
  // Clear debug logs from screen
  clearDebugLogs();
  
  // Clear chat messages
  const chatMessages = document.getElementById('chatMessages');
  if (chatMessages) {
    chatMessages.innerHTML = '';
  }
  
  // Add welcome message back if agent is active
  if (currentAgentName) {
    const welcomeMessage = `
      <div class="message bot-message welcome-message">
        <div class="message-avatar">
          <svg width="32" height="32" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
            <path d="M12 2C13.1 2 14 2.9 14 4C14 5.1 13.1 6 12 6C10.9 6 10 5.1 10 4C10 2.9 10.9 2 12 2ZM21 9V7L15 1H5C3.89 1 3 1.89 3 3V19C3 20.1 3.9 21 5 21H11V19H5V3H13V9H21Z" fill="currentColor"/>
          </svg>
        </div>
        <div class="message-content">
          <h3>Hello! I'm ${currentAgentName}</h3>
          <p>I'm your AI assistant for this Teams channel. I can help you with SharePoint documents and answer questions about your team's content.</p>
        </div>
      </div>
    `;
    
    if (chatMessages) {
      chatMessages.innerHTML = welcomeMessage;
    }
  }
  
  console.log('Chat and debug logs cleared');
}

// Function to safely initialize Teams
async function initializeTeams() {
    return new Promise((resolve, reject) => {
        showDebugMessage('Attempting to initialize Teams SDK...');
        
        // Check if Teams SDK is loaded
        if (typeof microsoftTeams === 'undefined') {
            const error = 'Microsoft Teams SDK is not loaded';
            showDebugMessage(error, true);
            reject(new Error(error));
            return;
        }
        
        // Set timeout for Teams initialization
        const timeout = setTimeout(() => {
            const error = 'Teams initialization timed out after 30 seconds';
            showDebugMessage(error, true);
            reject(new Error(error));
        }, 30000);
        
        try {
            microsoftTeams.app.initialize().then(() => {
                clearTimeout(timeout);
                showDebugMessage('Teams SDK initialized successfully');
                
                // Get Teams context with timeout
                const contextTimeout = setTimeout(() => {
                    const error = 'Getting Teams context timed out after 30 seconds';
                    showDebugMessage(error, true);
                    reject(new Error(error));
                }, 30000);
                
                microsoftTeams.app.getContext()
                    .then(context => {
                        clearTimeout(contextTimeout);
                        showDebugMessage('Teams context retrieved successfully');
                        resolve(context);
                    })
                    .catch(error => {
                        clearTimeout(contextTimeout);
                        const errorMsg = `Failed to get Teams context: ${error.message}`;
                        showDebugMessage(errorMsg, true);
                        reject(new Error(errorMsg));
                    });
            }).catch(error => {
                clearTimeout(timeout);
                const errorMsg = `Teams SDK initialization failed: ${error.message}`;
                showDebugMessage(errorMsg, true);
                reject(new Error(errorMsg));
            });
        } catch (error) {
            clearTimeout(timeout);
            const errorMsg = `Unexpected error during Teams initialization: ${error.message}`;
            showDebugMessage(errorMsg, true);
            reject(new Error(errorMsg));
        }
    });
}

function init() {
    showDebugMessage('Starting application initialization...');
    
    // Make sure all required elements exist
    const requiredElements = ['loadingScreen', 'firstScreen', 'secondScreen', 'thirdScreen', 'fourthScreen', 'fifthScreen'];
    const missingElements = requiredElements.filter(id => !document.getElementById(id));
    
    if (missingElements.length > 0) {
      console.error(`Missing required elements: ${missingElements.join(', ')}`);
      alert(`Error: Missing UI elements: ${missingElements.join(', ')}. Please refresh the page.`);
      return;
    }
    
    // Show the first screen
    document.getElementById('loadingScreen').style.display = 'none';
    showFirstScreen();
    
    // Initialize bot creation functionality
    initializeBotCreation();
    
    showDebugMessage('AI Agent Builder screen displayed as first screen');
}

// Check if the DOM is already loaded
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', init);
} else {
  // DOM is already loaded, run immediately
  setTimeout(init, 0);
}

// Clear any existing debug logs immediately
setTimeout(() => {
  clearDebugLogs();
}, 100);

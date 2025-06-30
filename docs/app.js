// Function to create agent by calling the Logic App
async function createAgent() {
    const agentName = document.getElementById('agentName').value.trim();
    const model = document.getElementById('modelSelect').value;
    
    if (!agentName) {
        alert('Please enter a name for your agent');
        return;
    }
    
    const createAgentBtn = document.getElementById('createAgentBtn');
    const originalText = createAgentBtn.textContent;
    
    try {
        // Disable button and show loading state
        createAgentBtn.disabled = true;
        createAgentBtn.textContent = 'Creating...';
        
        // Call the Logic App to create the agent
        const url = "https://prod-41.westus.logic.azure.com:443/workflows/e5f0ce23f3ea415696da0d9b4eeed2ec/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=IZXxoQiXyN8FToQ0GSaFPAy8iO9NEDf9vx5qRP7g0NA";
        
        const requestBody = {
            botName: agentName,
            botModel: model
        };
        
        const response = await fetch(url, {
            method: "POST",
            headers: {
                "Content-Type": "application/json"
            },
            body: JSON.stringify(requestBody)
        });
        
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        
        const data = await response.json();
        console.log("Flow response:", data);
        alert('Agent created successfully!');
        
    } catch (error) {
        console.error("Error creating agent:", error);
        alert(`Failed to create agent: ${error.message}`);
    } finally {
        // Re-enable button and restore text
        createAgentBtn.disabled = false;
        createAgentBtn.textContent = originalText;
    }
}

// Initialize the application when the page loads
document.addEventListener('DOMContentLoaded', function() {
    // Add click event for create agent button
    document.getElementById('createAgentBtn').addEventListener('click', createAgent);
    
    // No need for login button functionality anymore
    document.getElementById('loginBtn').style.display = 'none';
    const createAgentBtn = document.getElementById('createAgentBtn');
    const agentNameInput = document.getElementById('agentName');
    const modelSelect = document.getElementById('modelSelect');


    // Add click event to the create agent button
    createAgentBtn.addEventListener('click', async function() {
        const agentName = agentNameInput.value.trim();
        const selectedModel = modelSelect.value;

        if (!agentName) {
            alert('Please enter a name for your AI agent');
            return;
        }

        try {
            createAgentBtn.disabled = true;
            createAgentBtn.textContent = 'Saving...';
            
            // Save to SharePoint
            await saveToSharePoint({
                name: agentName,
                model: selectedModel
            });
            
            alert(`Agent "${agentName}" has been created successfully!`);
            
            // Reset form
            agentNameInput.value = '';
            modelSelect.value = 'gpt-4o';
        } catch (error) {
            console.error('Error:', error);
            alert('Failed to save agent. Please try again.');
        } finally {
            createAgentBtn.disabled = false;
            createAgentBtn.textContent = 'Create My Agent';
        }
    });

    // Add animation to the button on hover
    createAgentBtn.addEventListener('mouseenter', function() {
        this.style.transform = 'translateY(-2px)';
        this.style.boxShadow = '0 6px 16px rgba(121, 80, 242, 0.2)';
    });

    createAgentBtn.addEventListener('mouseleave', function() {
        this.style.transform = 'translateY(0)';
        this.style.boxShadow = 'none';
    });
});

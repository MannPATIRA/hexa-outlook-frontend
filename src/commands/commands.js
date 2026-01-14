/**
 * Commands for Ribbon Button Functions
 * These functions are triggered by ribbon button clicks
 */

// API Configuration (duplicated here for standalone usage)
// Always use deployed backend for both local and production
const getApiBaseUrl = () => {
    const stored = localStorage.getItem('procurement_api_url');
    if (stored) return stored;
    
    // Always use deployed backend
    return 'https://hexa-outlook-backend.onrender.com';
};
const API_BASE_URL = getApiBaseUrl();

Office.onReady((info) => {
    // Office is ready
    console.log('Commands: Office.js ready');
});

/**
 * Classify the current email from the ribbon button
 * @param {Office.AddinCommands.Event} event 
 */
async function classifyCurrentEmail(event) {
    try {
        const item = Office.context.mailbox.item;
        
        if (!item) {
            showNotification('Error', 'No email selected');
            event.completed();
            return;
        }

        // Get email details
        const subject = item.subject;
        
        // Get body asynchronously
        item.body.getAsync(Office.CoercionType.Text, async (bodyResult) => {
            if (bodyResult.status !== Office.AsyncResultStatus.Succeeded) {
                showNotification('Error', 'Failed to read email body');
                event.completed();
                return;
            }

            const body = bodyResult.value;
            const fromEmail = item.from ? item.from.emailAddress : 'unknown';

            try {
                // Call the backend API
                const response = await fetch(`${API_BASE_URL}/api/emails/classify`, {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json'
                    },
                    body: JSON.stringify({
                        email_chain: [{
                            subject: subject,
                            body: body,
                            from_email: fromEmail,
                            date: new Date().toISOString()
                        }],
                        most_recent_reply: {
                            subject: subject,
                            body: body,
                            from_email: fromEmail,
                            date: new Date().toISOString()
                        }
                    })
                });

                if (!response.ok) {
                    throw new Error(`API error: ${response.status}`);
                }

                const result = await response.json();
                
                // Show classification result
                const classification = result.classification || 'unknown';
                const confidence = result.confidence ? Math.round(result.confidence * 100) : 0;
                
                showNotification(
                    'Email Classified',
                    `Classification: ${classification.toUpperCase()}\nConfidence: ${confidence}%`
                );

            } catch (apiError) {
                console.error('API Error:', apiError);
                showNotification('Error', 'Failed to classify email. Is the backend running?');
            }

            event.completed();
        });

    } catch (error) {
        console.error('Error in classifyCurrentEmail:', error);
        showNotification('Error', 'An unexpected error occurred');
        event.completed();
    }
}

/**
 * Show a notification to the user
 */
function showNotification(title, message) {
    if (Office.context.mailbox.item) {
        Office.context.mailbox.item.notificationMessages.replaceAsync(
            'procurement-notification',
            {
                type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
                message: `${title}: ${message}`,
                icon: 'Icon.80x80',
                persistent: false
            }
        );
    }
}

// Register the function with Office
Office.actions = Office.actions || {};
Office.actions.associate("classifyCurrentEmail", classifyCurrentEmail);

// Also make it available globally
if (typeof window !== 'undefined') {
    window.classifyCurrentEmail = classifyCurrentEmail;
}

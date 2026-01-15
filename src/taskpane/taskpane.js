/**
 * Procurement Workflow Add-in - Main Taskpane Script
 * Handles all UI interactions and orchestrates the workflow
 */

// ==================== STATE MANAGEMENT ====================
const AppState = {
    // Currently loaded PRs
    prs: [],
    // Selected PR
    selectedPR: null,
    // Suppliers for selected PR
    suppliers: [],
    // Selected suppliers
    selectedSuppliers: [],
    // Generated RFQs
    rfqs: [],
    // Current email being processed
    currentEmail: null,
    // Email classification result
    classification: null,
    // Email processing result
    processingResult: null,
    // Available RFQs for quote comparison
    availableRfqs: [],
    // Current context mode (rfq-workflow, draft, clarification, quote)
    currentMode: 'rfq-workflow',
    // Current email context details
    emailContext: null,
    // Parsed questions with AI responses
    questions: [],
    // Pending RFQ drafts (for modal display)
    pendingDrafts: []
};

// ==================== STATE PERSISTENCE ====================
const STATE_KEY = 'procurement_addin_state';

function persistState(state) {
    try {
        const existing = getPersistedState();
        const merged = { ...existing, ...state, timestamp: Date.now() };
        localStorage.setItem(STATE_KEY, JSON.stringify(merged));
        console.log('State persisted:', merged);
    } catch (e) {
        console.error('Failed to persist state:', e);
    }
}

function getPersistedState() {
    try {
        const stored = localStorage.getItem(STATE_KEY);
        return stored ? JSON.parse(stored) : {};
    } catch (e) {
        console.error('Failed to get persisted state:', e);
        return {};
    }
}

function clearPersistedState() {
    try {
        localStorage.removeItem(STATE_KEY);
        console.log('Persisted state cleared');
    } catch (e) {
        console.error('Failed to clear persisted state:', e);
    }
}

// ==================== EMAIL ID MAPPING ====================
// Maps Outlook message IDs to backend email_ids for API calls
const EMAIL_ID_MAPPING_KEY = 'procurement_email_id_mapping';

/**
 * Store backend email_id for an Outlook message ID
 * The backend returns an email_id from /classify that must be used for all subsequent API calls
 */
function storeEmailId(outlookMessageId, backendEmailId) {
    try {
        const mapping = JSON.parse(localStorage.getItem(EMAIL_ID_MAPPING_KEY) || '{}');
        mapping[outlookMessageId] = backendEmailId;
        localStorage.setItem(EMAIL_ID_MAPPING_KEY, JSON.stringify(mapping));
        console.log(`Stored email_id mapping: ${outlookMessageId} -> ${backendEmailId}`);
    } catch (e) {
        console.error('Failed to store email_id:', e);
    }
}

/**
 * Retrieve stored backend email_id for an Outlook message ID
 */
function getStoredEmailId(outlookMessageId) {
    try {
        const mapping = JSON.parse(localStorage.getItem(EMAIL_ID_MAPPING_KEY) || '{}');
        return mapping[outlookMessageId] || null;
    } catch (e) {
        console.error('Failed to get stored email_id:', e);
        return null;
    }
}

/**
 * Store clarification_id for an email (needed for /suggest-response and /forward-to-engineering)
 */
function storeClarificationId(outlookMessageId, clarificationId) {
    try {
        const mapping = JSON.parse(localStorage.getItem(EMAIL_ID_MAPPING_KEY + '_clarifications') || '{}');
        mapping[outlookMessageId] = clarificationId;
        localStorage.setItem(EMAIL_ID_MAPPING_KEY + '_clarifications', JSON.stringify(mapping));
        console.log(`Stored clarification_id mapping: ${outlookMessageId} -> ${clarificationId}`);
    } catch (e) {
        console.error('Failed to store clarification_id:', e);
    }
}

/**
 * Retrieve stored clarification_id for an Outlook message ID
 */
function getStoredClarificationId(outlookMessageId) {
    try {
        const mapping = JSON.parse(localStorage.getItem(EMAIL_ID_MAPPING_KEY + '_clarifications') || '{}');
        return mapping[outlookMessageId] || null;
    } catch (e) {
        console.error('Failed to get stored clarification_id:', e);
        return null;
    }
}

/**
 * Ensure an email has been classified by the backend
 * If not already classified, calls /api/emails/classify first
 * Returns the backend email_id needed for subsequent API calls
 * 
 * @param {Object} email - The email object with id, subject, body, from, etc.
 * @param {string} expectedClassification - Expected type: 'quote' or 'clarification_request'
 * @returns {Object} - { emailId, classification, supplierId, rfqId }
 */
async function ensureEmailClassified(email, expectedClassification) {
    if (!email || !email.id) {
        throw new Error('Invalid email object');
    }
    
    // Check if we already have a stored backend email_id
    let backendEmailId = getStoredEmailId(email.id);
    
    if (backendEmailId) {
        console.log(`Email already classified, backend email_id: ${backendEmailId}`);
        return {
            emailId: backendEmailId,
            classification: expectedClassification,
            supplierId: email.from?.emailAddress?.address || 'unknown',
            rfqId: EmailOperations.extractRfqId ? EmailOperations.extractRfqId(email.subject) : null
        };
    }
    
    console.log('Email not yet classified, calling /api/emails/classify...');
    
    // Build email chain for classification
    const emailChain = [];
    
    // Try to get conversation emails if available
    if (email.conversationId && AuthService.isSignedIn()) {
        try {
            const escapedConvId = email.conversationId.replace(/'/g, "''");
            const response = await AuthService.graphRequest(
                `/me/messages?$filter=conversationId eq '${escapedConvId}'&$select=id,subject,from,body,receivedDateTime&$top=20`
            );
            
            if (response.value && response.value.length > 0) {
                // Sort by date ascending (oldest first)
                response.value.sort((a, b) => 
                    new Date(a.receivedDateTime) - new Date(b.receivedDateTime)
                );
                
                for (const convEmail of response.value) {
                    emailChain.push({
                        subject: convEmail.subject || '',
                        body: convEmail.body?.content || '',
                        from_email: convEmail.from?.emailAddress?.address || '',
                        date: convEmail.receivedDateTime || new Date().toISOString()
                    });
                }
            }
        } catch (convError) {
            console.warn('Failed to get conversation emails:', convError.message);
        }
    }
    
    // If no chain built, use just this email
    if (emailChain.length === 0) {
        emailChain.push({
            subject: email.subject || '',
            body: email.body?.content || email.bodyPreview || '',
            from_email: email.from?.emailAddress?.address || '',
            date: email.receivedDateTime || new Date().toISOString()
        });
    }
    
    // The most recent reply (the supplier's email)
    const mostRecentReply = {
        subject: email.subject || '',
        body: email.body?.content || email.bodyPreview || '',
        from_email: email.from?.emailAddress?.address || '',
        date: email.receivedDateTime || new Date().toISOString()
    };
    
    // Extract RFQ ID from subject
    const rfqId = EmailOperations.extractRfqId ? 
        EmailOperations.extractRfqId(email.subject) : 
        (email.subject?.match(/MAT-\d+/)?.[0] || null);
    
    // Supplier ID is the sender's email
    const supplierId = email.from?.emailAddress?.address || 'unknown';
    
    try {
        // Call the classify API
        const classifyResult = await ApiClient.classifyEmail(
            emailChain,
            mostRecentReply,
            rfqId,
            supplierId
        );
        
        console.log('Classification result:', classifyResult);
        
        // Store the backend email_id for future use
        if (classifyResult.email_id) {
            storeEmailId(email.id, classifyResult.email_id);
            backendEmailId = classifyResult.email_id;
        } else {
            // If backend doesn't return email_id, use outlook message id as fallback
            console.warn('Backend did not return email_id, using Outlook message ID');
            backendEmailId = email.id;
        }
        
        return {
            emailId: backendEmailId,
            classification: classifyResult.classification || expectedClassification,
            confidence: classifyResult.confidence,
            supplierId: supplierId,
            rfqId: rfqId
        };
    } catch (classifyError) {
        console.error('Classification API failed:', classifyError);
        // Return fallback using outlook message id
        return {
            emailId: email.id,
            classification: expectedClassification,
            supplierId: supplierId,
            rfqId: rfqId,
            error: classifyError.message
        };
    }
}

function restorePersistedState() {
    const state = getPersistedState();
    
    // Check if we just completed sending (add-in reopened after sending current draft)
    if (state.showSuccessOnReopen || state.lastSendResult === 'success') {
        const sent = state.sentCount || 0;
        const autoReplies = state.autoRepliesScheduled || 0;
        
        // Show success banner
        const successMessage = autoReplies > 0 
            ? `✓ Sent ${sent} RFQ(s) successfully! ${autoReplies} auto-replies scheduled - watch your inbox!`
            : `✓ Sent ${sent} RFQ(s) successfully!`;
        
        Helpers.showSuccess(successMessage);
        console.log('Restored state:', successMessage);
        
        // Clear all state
        clearPersistedState();
        return true; // Indicate state was restored
    }
    
    // Check if we were in the middle of sending (interrupted)
    if (state.sendingInProgress) {
        const sent = state.sentCount || 0;
        const total = state.totalDrafts || 0;
        
        if (sent > 0) {
            Helpers.showSuccess(`Sent ${sent}/${total} RFQ(s) before interruption. Please check your Sent RFQs folder.`);
        } else {
            Helpers.showError('Sending was interrupted. Please try again.');
        }
        clearPersistedState();
        return true;
    }
    
    // Handle partial success
    if (state.lastSendResult === 'partial') {
        Helpers.showSuccess(`Most RFQs sent successfully. Some may need to be resent.`);
        clearPersistedState();
        return true;
    }
    
    // Restore workflow state if recent (within last hour)
    // Don't restore selectedPR - let user select manually on page load
    const oneHourAgo = Date.now() - (60 * 60 * 1000);
    if (state.timestamp && state.timestamp > oneHourAgo) {
        // Skip restoring selectedPR to prevent auto-selection
        // AppState.selectedPR remains null until user explicitly selects one
        if (state.rfqs && state.rfqs.length > 0) {
            AppState.rfqs = state.rfqs;
            console.log('Restored RFQs:', state.rfqs.length);
        }
        if (state.currentStep) {
            console.log('Restored to step:', state.currentStep);
        }
    }
    
    return false;
}

// ==================== CONTEXT DETECTION ====================
/**
 * Detect the current email context to determine which UI mode to show
 * Returns: { type: 'draft'|'clarification'|'quote'|'normal'|'no-email', email?, item? }
 */
/**
 * Get all replies in a conversation, excluding emails from the current user
 * @param {string} conversationId - The conversation ID
 * @param {string} userEmail - The current user's email to exclude
 * @returns {Array} Array of reply emails, sorted by date descending (most recent first)
 */
async function getConversationReplies(conversationId, userEmail) {
    if (!conversationId || !AuthService.isSignedIn()) {
        return [];
    }
    
    try {
        console.log('Fetching conversation replies for:', conversationId);
        
        // Escape special characters in conversationId for OData filter
        const escapedConvId = conversationId.replace(/'/g, "''");
        
        // Fetch all emails in this conversation
        // Note: Personal Outlook accounts don't support $filter + $orderby together
        // so we fetch without orderby and sort in JavaScript
        const response = await AuthService.graphRequest(
            `/me/messages?$filter=conversationId eq '${escapedConvId}'&$select=id,subject,from,parentFolderId,categories,body,receivedDateTime,conversationId&$top=50`
        );
        
        if (!response.value || response.value.length === 0) {
            console.log('No emails found in conversation');
            return [];
        }
        
        console.log(`Found ${response.value.length} emails in conversation`);
        
        // Filter out emails from the current user (only keep supplier replies)
        const userEmailLower = userEmail.toLowerCase();
        const replies = response.value.filter(email => {
            const fromAddress = email.from?.emailAddress?.address?.toLowerCase() || '';
            return fromAddress !== userEmailLower;
        });
        
        // Sort by receivedDateTime descending (most recent first)
        replies.sort((a, b) => new Date(b.receivedDateTime) - new Date(a.receivedDateTime));
        
        console.log(`Found ${replies.length} replies from other senders`);
        return replies;
        
    } catch (error) {
        console.error('Error fetching conversation replies:', error);
        return [];
    }
}

/**
 * Check if an email is classified as a quote (via category or folder)
 */
function isQuoteEmail(email) {
    // Check categories
    const categories = email.categories || [];
    if (categories.some(c => c && c.toUpperCase().includes('QUOTE'))) {
        return true;
    }
    return false;
}

/**
 * Check if an email is classified as a clarification (via category or folder)
 */
function isClarificationEmail(email) {
    // Check categories
    const categories = email.categories || [];
    if (categories.some(c => c && (
        c.toUpperCase().includes('CLARIFICATION') || 
        c.toUpperCase().includes('YELLOW')
    ))) {
        return true;
    }
    return false;
}

/**
 * Check if an email is a Sent RFQ (via category or folder path)
 */
function isSentRfqEmail(email, folderPath) {
    // Check categories for "SENT RFQ"
    const categories = email.categories || [];
    if (categories.some(c => c && c.toUpperCase().includes('SENT RFQ'))) {
        return true;
    }
    
    // Check folder path
    if (folderPath) {
        const lowerPath = folderPath.toLowerCase();
        if (lowerPath.includes('sent rfq')) {
            return true;
        }
    }
    
    return false;
}

async function detectEmailContext() {
    console.log('=== Detecting email context ===');
    
    // Check if we have an email selected
    if (!Office.context || !Office.context.mailbox) {
        console.log('No Office context available');
        return { type: 'normal' };
    }

    const item = Office.context.mailbox.item;
    
    if (!item) {
        console.log('No email item selected');
        return { type: 'normal' };
    }

    // Check if we're in compose mode (draft)
    try {
        const itemType = item.itemType;
        console.log('Item type:', itemType);
        
        // Check if this is a compose (draft) context
        // In compose mode, item.body.setAsync exists (for writing)
        if (item.body && typeof item.body.setAsync === 'function') {
            console.log('Detected compose mode (setAsync available)');
            
            // Get subject - in compose mode this might be async or sync depending on version
            let subject = '';
            if (typeof item.subject === 'string') {
                subject = item.subject;
            } else if (item.subject && typeof item.subject.getAsync === 'function') {
                // Async subject access
                subject = await new Promise((resolve) => {
                    item.subject.getAsync((result) => {
                        resolve(result.status === Office.AsyncResultStatus.Succeeded ? result.value : '');
                    });
                });
            }
            
            console.log('Draft subject:', subject);
            
            // Check if this is an RFQ draft
            if (subject && subject.includes('RFQ')) {
                console.log('>>> Detected RFQ draft mode');
                return { type: 'draft', item: item };
            } else {
                // Still a draft, but not an RFQ - show draft mode anyway for any compose
                console.log('>>> Detected non-RFQ draft/compose mode');
                return { type: 'draft', item: item };
            }
        }
    } catch (e) {
        console.log('Error checking compose mode:', e);
    }

    // For read mode, check categories and folder
    if (!AuthService.isSignedIn()) {
        console.log('Not signed in - showing normal mode');
        return { type: 'normal', item: item };
    }

    try {
        // Get email ID - may need to convert REST ID
        let emailId = item.itemId;
        
        if (!emailId) {
            console.log('No email ID available - showing normal mode');
            return { type: 'normal', item: item };
        }

        console.log('Email ID (raw):', emailId.substring(0, 50) + '...');

        // Try to get email details (including conversationId for thread detection)
        let email = null;
        try {
            email = await AuthService.graphRequest(
                `/me/messages/${emailId}?$select=id,subject,from,parentFolderId,categories,body,receivedDateTime,conversationId`
            );
        } catch (graphError) {
            console.error('Graph API error getting email:', graphError);
            // Try converting the ID format
            if (Office.context.mailbox.convertToRestId) {
                try {
                    const restId = Office.context.mailbox.convertToRestId(
                        emailId, 
                        Office.MailboxEnums.RestVersion.v2_0
                    );
                    console.log('Converted to REST ID:', restId.substring(0, 50) + '...');
                    email = await AuthService.graphRequest(
                        `/me/messages/${restId}?$select=id,subject,from,parentFolderId,categories,body,receivedDateTime,conversationId`
                    );
                } catch (convertError) {
                    console.error('Error with converted ID:', convertError);
                }
            }
        }

        if (!email) {
            console.log('Could not fetch email details - showing normal mode');
            return { type: 'normal', item: item };
        }

        console.log('Email details:', {
            subject: email.subject,
            categories: email.categories,
            parentFolderId: email.parentFolderId ? email.parentFolderId.substring(0, 30) + '...' : null
        });

        // Check categories for classification
        const categories = email.categories || [];
        
        // Quote detection
        if (categories.some(c => c && c.toUpperCase().includes('QUOTE'))) {
            console.log('>>> Detected QUOTE email via category');
            return { type: 'quote', email: email, item: item };
        }
        
        // Clarification detection
        if (categories.some(c => c && (
            c.toUpperCase().includes('CLARIFICATION') || 
            c.toUpperCase().includes('YELLOW')
        ))) {
            console.log('>>> Detected CLARIFICATION email via category');
            return { type: 'clarification', email: email, item: item };
        }

        // Check folder - get both the immediate folder name AND the full path
        let folderPath = '';
        let folderName = '';
        const userEmail = Office.context.mailbox.userProfile?.emailAddress?.toLowerCase() || '';
        const emailFrom = email.from?.emailAddress?.address?.toLowerCase() || '';
        const isOwnEmail = userEmail && emailFrom && emailFrom === userEmail;
        
        console.log('=== Folder Detection ===');
        console.log('User email:', userEmail);
        console.log('Email from:', emailFrom);
        console.log('Is own email:', isOwnEmail);
        console.log('Conversation ID:', email.conversationId ? 'present' : 'MISSING');
        
        if (email.parentFolderId) {
            try {
                // First, get the immediate folder name directly (more reliable)
                const folderInfo = await AuthService.graphRequest(
                    `/me/mailFolders/${email.parentFolderId}?$select=displayName`
                );
                folderName = folderInfo?.displayName?.toLowerCase() || '';
                console.log('Immediate folder name:', folderName);
                
                // Also try to get full path for deeper matching
                try {
                    folderPath = await FolderManagement.getFolderPath(email.parentFolderId);
                    console.log('Full folder path:', folderPath);
                } catch (pathErr) {
                    console.log('Could not get full path, using folder name only');
                }
            } catch (folderErr) {
                console.error('Error getting folder info:', folderErr);
            }
        }
        
        // Check for Quotes folder (by immediate name OR full path)
        const lowerPath = folderPath.toLowerCase();
        const isInQuotesFolder = folderName.includes('quote') || lowerPath.includes('quote');
        const isInClarificationFolder = folderName.includes('clarification') || lowerPath.includes('clarification');
        const isInEngineerFolder = folderName.includes('engineer') || lowerPath.includes('engineer');
        
        console.log('Folder detection results:', { isInQuotesFolder, isInClarificationFolder, isInEngineerFolder });
        
        if (isInQuotesFolder) {
            console.log('>>> IN QUOTES FOLDER - switching to Quote mode');
            
            // If this is our own sent email, try to find the actual supplier reply
            if (isOwnEmail && email.conversationId) {
                console.log('>>> This is user sent email - looking for supplier quote reply...');
                try {
                    const replies = await getConversationReplies(email.conversationId, userEmail);
                    console.log('>>> Found', replies.length, 'supplier replies');
                    if (replies.length > 0) {
                        console.log('>>> Using supplier reply for Quote mode');
                        return { type: 'quote', email: replies[0], item: item, originalRfq: email };
                    }
                } catch (replyErr) {
                    console.error('Error getting replies:', replyErr);
                }
            }
            
            // Return quote mode with current email
            console.log('>>> Returning Quote mode with current email');
            return { type: 'quote', email: email, item: item };
        }
        
        if (isInClarificationFolder) {
            console.log('>>> IN CLARIFICATION FOLDER - switching to Clarification mode');
            
            // If this is our own sent email, try to find the actual supplier reply
            if (isOwnEmail && email.conversationId) {
                console.log('>>> This is user sent email - looking for supplier clarification...');
                try {
                    const replies = await getConversationReplies(email.conversationId, userEmail);
                    console.log('>>> Found', replies.length, 'supplier replies');
                    if (replies.length > 0) {
                        console.log('>>> Using supplier reply for Clarification mode');
                        return { type: 'clarification', email: replies[0], item: item, originalRfq: email };
                    }
                } catch (replyErr) {
                    console.error('Error getting replies:', replyErr);
                }
            }
            
            console.log('>>> Returning Clarification mode with current email');
            return { type: 'clarification', email: email, item: item };
        }
        
        if (isInEngineerFolder) {
            console.log('>>> IN ENGINEER FOLDER - switching to Clarification mode');
            return { type: 'clarification', email: email, item: item };
        }

        // ============================================================
        // SMART CONVERSATION DETECTION FOR SENT RFQs
        // If this is a sent RFQ, check if there are replies in the conversation
        // and show the appropriate mode based on the reply classification
        // ============================================================
        if (isSentRfqEmail(email, folderPath)) {
            console.log('>>> Detected SENT RFQ email - checking for conversation replies...');
            
            // Get the current user's email address
            const userEmail = Office.context.mailbox.userProfile?.emailAddress || '';
            
            if (email.conversationId && userEmail) {
                try {
                    const replies = await getConversationReplies(email.conversationId, userEmail);
                    
                    if (replies.length > 0) {
                        console.log(`Found ${replies.length} supplier replies in conversation`);
                        
                        // Check each reply for classification (most recent first)
                        for (const reply of replies) {
                            // First check if it's classified as a quote
                            if (isQuoteEmail(reply)) {
                                console.log('>>> Found QUOTE reply in conversation - switching to Quote mode');
                                return { 
                                    type: 'quote', 
                                    email: reply, 
                                    item: item, 
                                    originalRfq: email 
                                };
                            }
                            
                            // Then check if it's classified as a clarification
                            if (isClarificationEmail(reply)) {
                                console.log('>>> Found CLARIFICATION reply in conversation - switching to Clarification mode');
                                return { 
                                    type: 'clarification', 
                                    email: reply, 
                                    item: item, 
                                    originalRfq: email 
                                };
                            }
                        }
                        
                        // Check folder location of replies as fallback
                        for (const reply of replies) {
                            if (reply.parentFolderId) {
                                try {
                                    const replyFolderPath = await FolderManagement.getFolderPath(reply.parentFolderId);
                                    const lowerReplyPath = replyFolderPath.toLowerCase();
                                    
                                    if (lowerReplyPath.includes('quote')) {
                                        console.log('>>> Found reply in Quotes folder - switching to Quote mode');
                                        return { 
                                            type: 'quote', 
                                            email: reply, 
                                            item: item, 
                                            originalRfq: email 
                                        };
                                    }
                                    
                                    if (lowerReplyPath.includes('clarification')) {
                                        console.log('>>> Found reply in Clarification folder - switching to Clarification mode');
                                        return { 
                                            type: 'clarification', 
                                            email: reply, 
                                            item: item, 
                                            originalRfq: email 
                                        };
                                    }
                                } catch (e) {
                                    console.log('Could not get reply folder path:', e);
                                }
                            }
                        }
                        
                        // If we have unclassified replies, show the most recent one as a quote
                        // (most supplier replies are quotes)
                        console.log('>>> Found unclassified reply - showing as Quote mode by default');
                        return { 
                            type: 'quote', 
                            email: replies[0], 
                            item: item, 
                            originalRfq: email 
                        };
                    } else {
                        console.log('No supplier replies found in conversation - showing RFQ workflow');
                    }
                } catch (convError) {
                    console.error('Error checking conversation replies:', convError);
                }
            }
        }

        console.log('No special context detected - showing normal mode');
        return { type: 'normal', email: email, item: item };

    } catch (error) {
        console.error('Error detecting email context:', error);
        console.error('Stack:', error.stack);
        return { type: 'normal', item: item };
    }
}

// ==================== MODE RENDERING ====================
/**
 * Hide all mode containers and show main content
 */
function hideAllModes() {
    // Hide all mode containers
    const modeContainers = ['draft-mode', 'clarification-mode', 'quote-mode'];
    modeContainers.forEach(id => {
        const el = document.getElementById(id);
        if (el) el.classList.add('hidden');
    });
    
    // Show main content
    const mainContent = document.getElementById('main-content');
    if (mainContent) mainContent.style.display = 'block';
}

/**
 * Show a specific mode and hide the normal workflow
 */
function showMode(modeId) {
    console.log('showMode called with:', modeId);
    
    try {
        // Hide main content
        const mainContent = document.getElementById('main-content');
        if (mainContent) mainContent.style.display = 'none';
        
        // Hide all mode containers first
        const modeContainers = ['draft-mode', 'clarification-mode', 'quote-mode'];
        modeContainers.forEach(id => {
            const el = document.getElementById(id);
            if (el) el.classList.add('hidden');
        });
        
        // Show the requested mode
        const modeEl = document.getElementById(modeId);
        if (modeEl) {
            modeEl.classList.remove('hidden');
            console.log('Mode element shown:', modeId);
        } else {
            console.error('Mode element not found:', modeId);
            // Fallback - show main content
            if (mainContent) mainContent.style.display = 'block';
        }
    } catch (error) {
        console.error('Error in showMode:', error);
        // Try to show main content as fallback
        const mainContent = document.getElementById('main-content');
        if (mainContent) mainContent.style.display = 'block';
    }
}

/**
 * Render UI based on detected email context
 */
async function renderContextUI(context) {
    console.log('Rendering context UI for type:', context?.type);
    
    try {
        AppState.emailContext = context;
        AppState.currentMode = context?.type || 'normal';
        
        switch (context?.type) {
            case 'draft':
                try {
                    await showDraftMode(context);
                } catch (e) {
                    console.error('Error showing draft mode:', e);
                    showRFQWorkflowMode();
                }
                break;
            case 'clarification':
                try {
                    await showClarificationMode(context);
                } catch (e) {
                    console.error('Error showing clarification mode:', e);
                    showRFQWorkflowMode();
                }
                break;
            case 'quote':
                try {
                    await showQuoteMode(context);
                } catch (e) {
                    console.error('Error showing quote mode:', e);
                    showRFQWorkflowMode();
                }
                break;
            case 'normal':
            case 'no-email':
            default:
                showRFQWorkflowMode();
                break;
        }
    } catch (error) {
        console.error('Error in renderContextUI:', error);
        showRFQWorkflowMode();
    }
}

/**
 * Show RFQ Workflow mode (normal mode)
 */
function showRFQWorkflowMode() {
    console.log('Showing RFQ Workflow mode');
    hideAllModes();
    AppState.currentMode = 'rfq-workflow';
    
    // Restore header title
    const headerTitle = document.getElementById('header-title');
    if (headerTitle) {
        headerTitle.textContent = '';
        headerTitle.removeAttribute('title');
    }
    
    // Show RFQ workflow section
    const mainContent = document.getElementById('main-content');
    const rfqWorkflowTab = document.getElementById('rfq-workflow-tab');
    if (mainContent && rfqWorkflowTab) {
        // Hide other tab sections
        document.querySelectorAll('.tab-content').forEach(tab => {
            tab.classList.remove('active');
            tab.classList.add('hidden');
        });
        
        // Show RFQ workflow
        rfqWorkflowTab.classList.remove('hidden');
        rfqWorkflowTab.classList.add('active');
        mainContent.style.display = 'block';
    }
}

/**
 * Show Draft mode when user is viewing a draft email
 */
async function showDraftMode(context) {
    console.log('Showing Draft mode');
    showMode('draft-mode');
    AppState.currentMode = 'draft';

    // Update header title
    const headerTitle = document.getElementById('header-title');
    if (headerTitle) {
        headerTitle.textContent = 'RFQ Draft Editor';
        headerTitle.setAttribute('title', 'RFQ Draft Editor');
    }

    // Load pending RFQ drafts
    await loadPendingDrafts();
}

/**
 * Extract summary data from drafts array
 * @param {Array} drafts - Array of draft email objects
 * @returns {Object} Summary object with count, material, quantity, supplierCount
 */
function extractDraftSummary(drafts) {
    if (!drafts || drafts.length === 0) {
        return { count: 0, material: null, quantity: null, supplierCount: 0 };
    }

    const materials = [];
    const quantities = [];
    const suppliers = new Set();

    drafts.forEach(draft => {
        const subject = draft.subject || '';
        
        // Extract material code
        const materialMatch = subject.match(/MAT-\d+/i);
        if (materialMatch) {
            materials.push(materialMatch[0]);
        }
        
        // Extract quantity (pattern: number followed by optional unit like "pcs", "units", etc.)
        // Look for patterns like " - 100 pcs" or "100 pcs" (avoid matching material codes like MAT-12345)
        const quantityMatch = subject.match(/[-–—]\s*(\d+)\s*(pcs|units|pieces)?/i) || subject.match(/\b(\d+)\s+(pcs|units|pieces)\b/i);
        if (quantityMatch) {
            quantities.push(quantityMatch[1] + (quantityMatch[2] ? ' ' + quantityMatch[2] : ''));
        }
        
        // Count unique suppliers
        const recipient = draft.toRecipients?.[0]?.emailAddress?.address;
        if (recipient) {
            suppliers.add(recipient);
        }
    });

    // Find most common material
    const materialCounts = {};
    materials.forEach(m => materialCounts[m] = (materialCounts[m] || 0) + 1);
    const mostCommonMaterial = Object.keys(materialCounts).reduce((a, b) => 
        materialCounts[a] > materialCounts[b] ? a : b, materials[0] || null
    );

    // Find most common quantity
    const quantityCounts = {};
    quantities.forEach(q => quantityCounts[q] = (quantityCounts[q] || 0) + 1);
    const mostCommonQuantity = Object.keys(quantityCounts).reduce((a, b) => 
        quantityCounts[a] > quantityCounts[b] ? a : b, quantities[0] || null
    );

    return {
        count: drafts.length,
        material: mostCommonMaterial || null,
        quantity: mostCommonQuantity || null,
        supplierCount: suppliers.size
    };
}

/**
 * Load and display pending RFQ drafts
 */
async function loadPendingDrafts() {
    const summaryCard = document.getElementById('draft-summary-card');
    const draftActionsSection = document.getElementById('draft-actions-section');
    const sendStatus = document.getElementById('draft-send-status');
    const progressTracker = document.getElementById('rfq-progress-tracker');
    
    if (!summaryCard) return;
    
    summaryCard.innerHTML = '<p class="loading-text">Loading drafts...</p>';
    
    // Hide progress elements (those are for during/after sending)
    if (sendStatus) sendStatus.classList.add('hidden');
    if (progressTracker) progressTracker.classList.add('hidden');
    
    if (!AuthService.isSignedIn()) {
        summaryCard.innerHTML = '<p class="loading-text">Please sign in to view drafts</p>';
        return;
    }
    
    try {
        // Try to get drafts - use a simpler query if the filter fails
        let drafts = null;
        try {
            drafts = await AuthService.graphRequest(
                `/me/mailFolders/Drafts/messages?$filter=startswith(subject,'RFQ for')&$select=id,subject,toRecipients,createdDateTime&$top=20&$orderby=createdDateTime desc`
            );
        } catch (filterError) {
            // If filter fails, try without filter (get all drafts)
            console.warn('Filter query failed, trying without filter:', filterError);
            try {
                drafts = await AuthService.graphRequest(
                    `/me/mailFolders/Drafts/messages?$select=id,subject,toRecipients,createdDateTime&$top=20&$orderby=createdDateTime desc`
                );
                // Filter client-side
                if (drafts.value) {
                    drafts.value = drafts.value.filter(d => 
                        d.subject && d.subject.toLowerCase().startsWith('rfq for')
                    );
                }
            } catch (simpleError) {
                console.error('Both draft queries failed:', simpleError);
                // Still show the send button - user is viewing a draft
                summaryCard.innerHTML = `
                    <div class="no-drafts-message">
                        <div class="icon"></div>
                        <p>You're viewing an RFQ draft. Click "Send all RFQs" to send all pending drafts.</p>
                    </div>
                `;
                if (draftActionsSection) draftActionsSection.classList.remove('hidden');
                const sendBtn = document.getElementById('send-all-drafts-btn');
                const viewDetailsBtn = document.getElementById('view-draft-details-btn');
                if (sendBtn) sendBtn.disabled = false;
                if (viewDetailsBtn) viewDetailsBtn.disabled = true;
                AppState.pendingDrafts = [];
                return;
            }
        }
        
        if (!drafts.value || drafts.value.length === 0) {
            summaryCard.innerHTML = `
                <div class="no-drafts-message">
                    <div class="icon"></div>
                    <p>No RFQ drafts found. Generate RFQs in the workflow first.</p>
                </div>
            `;
            if (draftActionsSection) draftActionsSection.classList.add('hidden');
            AppState.pendingDrafts = [];
            return;
        }
        
        // Store drafts for modal access
        AppState.pendingDrafts = drafts.value;
        
        // Show draft actions
        if (draftActionsSection) draftActionsSection.classList.remove('hidden');
        
        // Extract summary data
        const summary = extractDraftSummary(drafts.value);
        
        // Render summary card
        let summaryHTML = `<div class="draft-summary-item">${summary.count} drafts prepared</div>`;
        if (summary.material) {
            summaryHTML += `<div class="draft-summary-item">Material ${Helpers.escapeHtml(summary.material)}</div>`;
        }
        if (summary.quantity) {
            summaryHTML += `<div class="draft-summary-item">Quantity ${Helpers.escapeHtml(summary.quantity)}</div>`;
        }
        summaryHTML += `<div class="draft-summary-item">Suppliers ${summary.supplierCount}</div>`;
        
        summaryCard.innerHTML = summaryHTML;
        
        // Enable buttons
        const sendBtn = document.getElementById('send-all-drafts-btn');
        const viewDetailsBtn = document.getElementById('view-draft-details-btn');
        if (sendBtn) sendBtn.disabled = false;
        if (viewDetailsBtn) viewDetailsBtn.disabled = false;
        
    } catch (error) {
        console.error('Error loading drafts:', error);
        // Don't show error - just show a helpful message
        summaryCard.innerHTML = `
            <div class="no-drafts-message">
                <div class="icon"></div>
                <p>You're viewing an RFQ draft. Click "Send all RFQs" to send all pending drafts.</p>
            </div>
        `;
        if (draftActionsSection) draftActionsSection.classList.remove('hidden');
        const sendBtn = document.getElementById('send-all-drafts-btn');
        const viewDetailsBtn = document.getElementById('view-draft-details-btn');
        if (sendBtn) sendBtn.disabled = false;
        if (viewDetailsBtn) viewDetailsBtn.disabled = true;
        AppState.pendingDrafts = [];
    }
}

/**
 * Show draft details modal with full list of drafts
 */
function showDraftDetailsModal() {
    const modal = document.getElementById('draft-details-modal');
    const listContainer = document.getElementById('draft-details-list');
    
    if (!modal || !listContainer) return;
    
    const drafts = AppState.pendingDrafts || [];
    
    if (drafts.length === 0) {
        listContainer.innerHTML = '<p class="placeholder-text">No drafts available</p>';
    } else {
        // Render the full draft list
        listContainer.innerHTML = drafts.map(draft => {
            const recipient = draft.toRecipients?.[0]?.emailAddress?.name || 
                             draft.toRecipients?.[0]?.emailAddress?.address || 
                             'Unknown';
            return `
                <div class="draft-item" data-draft-id="${draft.id}">
                    <div class="draft-item-info">
                        <div class="draft-item-supplier">${Helpers.escapeHtml(recipient)}</div>
                        <div class="draft-item-subject">${Helpers.escapeHtml(draft.subject)}</div>
                    </div>
                    <span class="draft-item-status">Draft</span>
                </div>
            `;
        }).join('');
    }
    
    modal.classList.remove('hidden');
}

/**
 * Close draft details modal
 */
function closeDraftDetailsModal() {
    const modal = document.getElementById('draft-details-modal');
    if (modal) {
        modal.classList.add('hidden');
    }
}

/**
 * Load and display RFQ progress (sent, auto-replies, sorted)
 */
async function loadRfqProgress(state) {
    const sentCount = state.sentCount || 0;
    const autoRepliesScheduled = state.autoRepliesScheduled || 0;
    const materialCodes = state.materialCodes || [];
    
    // Warn if no material codes tracked (legacy state)
    if (materialCodes.length === 0 && sentCount > 0) {
        console.warn('No material codes in state - counting all folders (may include old replies from previous RFQs)');
    } else if (materialCodes.length > 0) {
        console.log(`Filtering replies for material codes: ${materialCodes.join(', ')}`);
    }
    
    // Helper function to set progress item state
    const setProgressItemState = (itemElement, state) => {
        if (!itemElement) return;
        itemElement.classList.remove('active', 'completed', 'not-started');
        itemElement.classList.add(state);
    };
    
    // Helper to check if email is an undeliverable/bounceback
    const isUndeliverable = (email, bodyPreview = '') => {
        const subject = (email.subject || '').toLowerCase();
        const from = (email.from?.emailAddress?.address || '').toLowerCase();
        const fromName = (email.from?.emailAddress?.name || '').toLowerCase();
        const body = (bodyPreview || '').toLowerCase();
        
        // Subject/from checks (existing)
        if (subject.includes('undeliverable') || 
            subject.includes('delivery failure') ||
            subject.includes('delivery has failed') ||
            subject.includes('mail delivery failed') ||
            from.includes('postmaster') ||
            from.includes('mailer-daemon') ||
            (from.includes('noreply') && subject.includes('failed')) ||
            fromName.includes('postmaster') ||
            fromName.includes('mailer-daemon')) {
            return true;
        }
        
        // NEW: Body content checks for bounceback patterns
        if (body.includes('message undeliverable') ||
            body.includes('delivery has failed') ||
            body.includes('returned mail') ||
            body.includes('mail delivery subsystem') ||
            body.includes('delivery status notification') ||
            body.includes('this is an automatically generated delivery status notification') ||
            body.includes('delivery to the following recipient failed') ||
            body.includes('could not be delivered')) {
            return true;
        }
        
        return false;
    };
    
    // Helper to check if email is a real supplier reply
    const isSupplierReply = (email, bodyPreview = '') => {
        const subject = (email.subject || '').toLowerCase();
        
        // Must contain RFQ in subject
        if (!subject.includes('rfq')) {
            return false;
        }
        
        // Must not be undeliverable
        if (isUndeliverable(email, bodyPreview)) {
            return false;
        }
        
        // NEW: Must have actual content (not just bounceback)
        const body = (bodyPreview || '').trim();
        if (body.length < 50) {
            return false; // Too short to be a real reply
        }
        
        // Check body doesn't contain bounceback patterns
        const bouncePatterns = [
            'delivery failed',
            'undeliverable',
            'returned mail',
            'mail delivery subsystem',
            'delivery status notification',
            'could not be delivered',
            'permanent failure',
            'temporary failure'
        ];
        if (bouncePatterns.some(pattern => body.includes(pattern))) {
            return false;
        }
        
        return true;
    };
    
    // Try to count actual replies received
    let repliesReceived = 0;
    let repliesSorted = 0;
    const allReplyIds = new Set(); // Track unique reply IDs
    
    if (AuthService.isSignedIn()) {
        try {
            // Step 1: Get all mail folders and find material code folders first
            let quoteFolders = [];
            let clarificationFolders = [];
            
            const allFolders = await AuthService.graphRequest(
                `/me/mailFolders?$select=id,displayName,parentFolderId&$top=500`
            );
            
            if (allFolders.value) {
                // First find material code folders (MAT-XXXXX pattern)
                // Only include folders for material codes that were sent in the current batch
                const materialFolders = [];
                for (const folder of allFolders.value) {
                    if (/^MAT-\d+$/i.test(folder.displayName)) {
                        const folderCode = folder.displayName.toUpperCase();
                        // Only include if materialCodes is empty (legacy/fallback) or if this material code was tracked
                        if (materialCodes.length === 0 || materialCodes.includes(folderCode)) {
                            materialFolders.push(folder);
                        }
                    }
                }
                
                // Then find Quotes/Clarification folders within each material folder
                for (const materialFolder of materialFolders) {
                    try {
                        const materialSubfolders = await AuthService.graphRequest(
                            `/me/mailFolders/${materialFolder.id}/childFolders?$select=id,displayName&$top=20`
                        );
                        
                        if (materialSubfolders.value) {
                            for (const subfolder of materialSubfolders.value) {
                                const name = (subfolder.displayName || '').toLowerCase();
                                if (name.includes('quote') && !name.includes('sent')) {
                                    quoteFolders.push(subfolder.id);
                                }
                                if (name.includes('clarification') && !name.includes('awaiting')) {
                                    clarificationFolders.push(subfolder.id);
                                }
                            }
                        }
                    } catch (e) {
                        // Ignore errors fetching subfolders for a specific material folder
                        console.warn(`Error fetching subfolders for ${materialFolder.displayName}:`, e);
                    }
                }
                
                // Also check for top-level folders as fallback (in case structure is different)
                for (const folder of allFolders.value) {
                    const folderName = (folder.displayName || '').toLowerCase();
                    if (folderName.includes('quote') && !folderName.includes('sent') && !quoteFolders.includes(folder.id)) {
                        quoteFolders.push(folder.id);
                    }
                    if (folderName.includes('clarification') && !folderName.includes('awaiting') && !clarificationFolders.includes(folder.id)) {
                        clarificationFolders.push(folder.id);
                    }
                }
            }
            
            // Step 2: Count emails in sorted folders (Quotes and Clarifications)
            for (const folderId of quoteFolders) {
                try {
                    const folderEmails = await AuthService.graphRequest(
                        `/me/mailFolders/${folderId}/messages?$select=id,subject,from,bodyPreview&$top=100`
                    );
                    if (folderEmails.value) {
                        for (const email of folderEmails.value) {
                            const bodyPreview = email.bodyPreview || '';
                            if (isSupplierReply(email, bodyPreview) && !allReplyIds.has(email.id)) {
                                allReplyIds.add(email.id);
                                repliesReceived++;
                                repliesSorted++;
                            }
                        }
                    }
                } catch (e) {
                    console.warn(`Error counting emails in quote folder ${folderId}:`, e);
                }
            }
            
            for (const folderId of clarificationFolders) {
                try {
                    const folderEmails = await AuthService.graphRequest(
                        `/me/mailFolders/${folderId}/messages?$select=id,subject,from,bodyPreview&$top=100`
                    );
                    if (folderEmails.value) {
                        for (const email of folderEmails.value) {
                            const bodyPreview = email.bodyPreview || '';
                            if (isSupplierReply(email, bodyPreview) && !allReplyIds.has(email.id)) {
                                allReplyIds.add(email.id);
                                repliesReceived++;
                                repliesSorted++;
                            }
                        }
                    }
                } catch (e) {
                    console.warn(`Error counting emails in clarification folder ${folderId}:`, e);
                }
            }
            
            // Step 3: Count RFQ-related emails in inbox (not yet sorted)
            try {
                const inboxReplies = await AuthService.graphRequest(
                    `/me/mailFolders/inbox/messages?$filter=contains(subject,'RFQ')&$top=100&$select=id,subject,from,bodyPreview&$orderby=receivedDateTime desc`
                );
                
                if (inboxReplies.value) {
                    for (const email of inboxReplies.value) {
                        const bodyPreview = email.bodyPreview || '';
                        if (isSupplierReply(email, bodyPreview) && !allReplyIds.has(email.id)) {
                            allReplyIds.add(email.id);
                            repliesReceived++;
                            // Not sorted yet, so don't increment repliesSorted
                        }
                    }
                }
            } catch (e) {
                console.warn('Error getting inbox replies:', e);
            }
            
        } catch (err) {
            console.warn('Error counting replies:', err);
        }
    }
    
    // Identify current active stage and format output
    const statusMessage = document.getElementById('rfq-status-message');
    
    // Stage 1: Sent
    const sentItem = document.getElementById('progress-item-sent');
    const sentCountEl = document.getElementById('sent-rfq-count');
    if (sentItem && sentCountEl) {
        if (sentCount > 0) {
            setProgressItemState(sentItem, 'completed');
            sentCountEl.textContent = '✓';
        } else {
            setProgressItemState(sentItem, 'active');
            sentCountEl.textContent = '0';
        }
    }

    // Stage 2: Scheduled
    const scheduledItem = document.getElementById('progress-item-scheduled');
    const scheduledCountEl = document.getElementById('auto-replies-scheduled-count');
    if (scheduledItem && scheduledCountEl) {
        if (sentCount > 0 && autoRepliesScheduled === sentCount) {
            setProgressItemState(scheduledItem, 'completed');
            scheduledCountEl.textContent = '✓';
        } else if (sentCount > 0) {
            setProgressItemState(scheduledItem, 'active');
            scheduledCountEl.textContent = `${autoRepliesScheduled} of ${sentCount}`;
        } else {
            setProgressItemState(scheduledItem, 'not-started');
            scheduledCountEl.textContent = 'not started';
        }
    }

    // Stage 3: Received
    const receivedItem = document.getElementById('progress-item-received');
    const receivedCountEl = document.getElementById('replies-received-count');
    const receivedProgress = document.getElementById('replies-received-progress');
    if (receivedItem && receivedCountEl) {
        if (sentCount > 0 && repliesReceived === sentCount) {
            setProgressItemState(receivedItem, 'completed');
            receivedCountEl.textContent = '✓';
            if (receivedProgress) receivedProgress.style.width = '100%';
        } else if (sentCount > 0 && autoRepliesScheduled === sentCount) {
            setProgressItemState(receivedItem, 'active');
            receivedCountEl.textContent = `${repliesReceived} of ${sentCount}`;
            if (receivedProgress) receivedProgress.style.width = `${(repliesReceived / sentCount) * 100}%`;
        } else {
            setProgressItemState(receivedItem, 'not-started');
            receivedCountEl.textContent = sentCount > 0 ? `0 of ${sentCount}` : 'not started';
            if (receivedProgress) receivedProgress.style.width = '0%';
        }
    }

    // Stage 4: Sorted
    const sortedItem = document.getElementById('progress-item-sorted');
    const sortedCountEl = document.getElementById('replies-sorted-count');
    const sortedProgress = document.getElementById('replies-sorted-progress');
    if (sortedItem && sortedCountEl) {
        if (repliesReceived > 0 && repliesSorted === repliesReceived) {
            setProgressItemState(sortedItem, 'completed');
            sortedCountEl.textContent = '✓';
            if (sortedProgress) sortedProgress.style.width = '100%';
        } else if (repliesReceived > 0) {
            setProgressItemState(sortedItem, 'active');
            sortedCountEl.textContent = `${repliesSorted} of ${repliesReceived}`;
            if (sortedProgress) sortedProgress.style.width = `${(repliesSorted / repliesReceived) * 100}%`;
        } else {
            setProgressItemState(sortedItem, 'not-started');
            sortedCountEl.textContent = 'not started';
            if (sortedProgress) sortedProgress.style.width = '0%';
        }
    }

    // Update state message
    if (statusMessage) {
        if (sentCount === 0) {
            statusMessage.textContent = 'Generating RFQs...';
        } else if (autoRepliesScheduled < sentCount) {
            statusMessage.textContent = 'Scheduling auto-replies for sent RFQs...';
        } else if (repliesReceived < sentCount) {
            statusMessage.textContent = 'Waiting for supplier replies. We will notify you as they arrive.';
        } else if (repliesSorted < repliesReceived) {
            statusMessage.textContent = 'Sorting received replies into your project folders...';
        } else {
            statusMessage.textContent = 'All RFQs sent and replies processed.';
        }
    }
}

/**
 * Render question cards with expandable functionality
 * @param {Array} questions - Array of question objects
 * @param {Object} email - Email context
 */
function renderQuestionCards(questions, email) {
    const questionsList = document.getElementById('clarification-questions-list');
    if (!questionsList) return;
    
    questionsList.innerHTML = '';
    
    // Group questions by category
    const questionsByCategory = {};
    questions.forEach(q => {
        const category = q.category || 'General Questions';
        if (!questionsByCategory[category]) {
            questionsByCategory[category] = [];
        }
        questionsByCategory[category].push(q);
    });
    
    let globalQuestionIndex = 0;
    
    Object.keys(questionsByCategory).forEach(category => {
        const categoryQuestions = questionsByCategory[category];
        
        // Add category header if multiple categories
        if (Object.keys(questionsByCategory).length > 1) {
            const categoryHeader = document.createElement('div');
            categoryHeader.className = 'question-category-header';
            categoryHeader.textContent = category;
            questionsList.appendChild(categoryHeader);
        }
        
        // Create card for each question
        categoryQuestions.forEach((q) => {
            globalQuestionIndex++;
            q.displayIndex = globalQuestionIndex;
            
            const questionCard = document.createElement('div');
            questionCard.className = 'question-card';
            questionCard.dataset.questionId = q.id;
            
            // Card header (always visible)
            const cardHeader = document.createElement('div');
            cardHeader.className = 'question-card-header';
            cardHeader.onclick = () => toggleQuestionCard(q.id);
            
            const headerContent = document.createElement('div');
            headerContent.className = 'question-card-header-content';
            
            const questionNumber = document.createElement('span');
            questionNumber.className = 'question-number';
            questionNumber.textContent = `${globalQuestionIndex}.`;
            
            const questionText = document.createElement('div');
            questionText.className = 'question-text';
            questionText.textContent = q.question;
            
            const expandIcon = document.createElement('span');
            expandIcon.className = 'question-expand-icon';
            expandIcon.innerHTML = q.isExpanded ? '▼' : '▶';
            
            headerContent.appendChild(questionNumber);
            headerContent.appendChild(questionText);
            cardHeader.appendChild(headerContent);
            cardHeader.appendChild(expandIcon);
            
            // Card content (expandable)
            const cardContent = document.createElement('div');
            cardContent.className = 'question-card-content';
            cardContent.style.display = q.isExpanded ? 'block' : 'none';
            
            // AI Response section
            const aiSection = document.createElement('div');
            aiSection.className = 'ai-response-section';
            
            const aiLabel = document.createElement('label');
            aiLabel.className = 'response-label';
            aiLabel.textContent = 'AI Response';
            
            const aiTextarea = document.createElement('textarea');
            aiTextarea.className = 'response-textarea ai-response-textarea';
            aiTextarea.rows = 4;
            aiTextarea.value = q.aiResponse || '';
            aiTextarea.placeholder = q.isLoadingResponse ? 'Generating...' : 'AI response';
            aiTextarea.disabled = q.isLoadingResponse;
            aiTextarea.oninput = (e) => {
                const question = AppState.questions.find(qq => qq.id === q.id);
                if (question) question.aiResponse = e.target.value;
            };
            
            const aiLoading = document.createElement('div');
            aiLoading.className = 'response-loading';
            aiLoading.style.display = q.isLoadingResponse ? 'flex' : 'none';
            aiLoading.innerHTML = '<div class="spinner-small"></div><span>Generating AI response...</span>';
            
            // "Enter custom response" button - only show when AI response exists and is not loading
            const dontLikeButton = document.createElement('button');
            dontLikeButton.type = 'button';
            dontLikeButton.className = 'dont-like-button';
            dontLikeButton.textContent = 'Enter custom response';
            dontLikeButton.style.display = (!q.isLoadingResponse && q.aiResponse && !q.showCustomResponse) ? 'flex' : 'none';
            dontLikeButton.onclick = () => {
                const question = AppState.questions.find(qq => qq.id === q.id);
                if (question) {
                    question.showCustomResponse = true;
                    question.useCustomResponse = true;
                    updateQuestionCard(q.id);
                }
            };
            
            aiSection.appendChild(aiLabel);
            aiSection.appendChild(aiTextarea);
            aiSection.appendChild(aiLoading);
            aiSection.appendChild(dontLikeButton);
            
            // Custom Response section - hidden by default
            const customSection = document.createElement('div');
            customSection.className = 'custom-response-section';
            customSection.style.display = q.showCustomResponse ? 'block' : 'none';
            
            const customLabel = document.createElement('label');
            customLabel.className = 'response-label';
            customLabel.textContent = 'Custom Response';
            
            const customTextarea = document.createElement('textarea');
            customTextarea.className = 'response-textarea custom-response-textarea';
            customTextarea.rows = 5;
            customTextarea.value = q.customResponse || '';
            customTextarea.placeholder = 'Enter your custom response here...';
            customTextarea.oninput = (e) => {
                const question = AppState.questions.find(qq => qq.id === q.id);
                if (question) {
                    question.customResponse = e.target.value;
                    // Automatically use custom response when user types
                    if (e.target.value.trim()) {
                        question.useCustomResponse = true;
                    }
                }
            };
            
            const useCustomCheckbox = document.createElement('label');
            useCustomCheckbox.className = 'use-custom-checkbox';
            const checkbox = document.createElement('input');
            checkbox.type = 'checkbox';
            checkbox.checked = q.useCustomResponse;
            checkbox.onchange = (e) => {
                const question = AppState.questions.find(qq => qq.id === q.id);
                if (question) question.useCustomResponse = e.target.checked;
            };
            useCustomCheckbox.appendChild(checkbox);
            useCustomCheckbox.appendChild(document.createTextNode(' Use custom response instead of AI response'));
            
            customSection.appendChild(customLabel);
            customSection.appendChild(customTextarea);
            customSection.appendChild(useCustomCheckbox);
            
            cardContent.appendChild(aiSection);
            cardContent.appendChild(customSection);
            
            questionCard.appendChild(cardHeader);
            questionCard.appendChild(cardContent);
            questionsList.appendChild(questionCard);
        });
    });
}

/**
 * Toggle question card expand/collapse
 */
function toggleQuestionCard(questionId) {
    const question = AppState.questions.find(q => q.id === questionId);
    if (!question) return;
    
    question.isExpanded = !question.isExpanded;
    
    const card = document.querySelector(`[data-question-id="${questionId}"]`);
    if (!card) return;
    
    const content = card.querySelector('.question-card-content');
    const icon = card.querySelector('.question-expand-icon');
    
    if (content) {
        content.style.display = question.isExpanded ? 'block' : 'none';
    }
    if (icon) {
        icon.innerHTML = question.isExpanded ? '▼' : '▶';
    }
}

/**
 * Generate AI responses for all questions
 */
async function generateAIResponsesForQuestions(questions, email) {
    const emailContext = {
        subject: email.subject || '',
        body: email.body?.content || '',
        rfqContext: '' // Can be enhanced with RFQ details if available
    };
    
    // Generate responses for each question sequentially to avoid rate limiting
    for (const question of questions) {
        question.isLoadingResponse = true;
        updateQuestionCard(question.id);
        
        try {
            const aiResponse = await OpenAIService.generateResponse(question.question, emailContext);
            question.aiResponse = aiResponse;
            question.hasError = false;
        } catch (error) {
            console.error(`Error generating response for question ${question.id}:`, error);
            question.hasError = true;
            question.aiResponse = 'Error generating response. Please provide a custom response.';
        } finally {
            question.isLoadingResponse = false;
            updateQuestionCard(question.id);
        }
    }
}

/**
 * Update a specific question card in the DOM
 */
function updateQuestionCard(questionId) {
    const question = AppState.questions.find(q => q.id === questionId);
    if (!question) return;
    
    const card = document.querySelector(`[data-question-id="${questionId}"]`);
    if (!card) return;
    
    const aiTextarea = card.querySelector('.ai-response-textarea');
    const aiLoading = card.querySelector('.response-loading');
    const dontLikeButton = card.querySelector('.dont-like-button');
    const customSection = card.querySelector('.custom-response-section');
    const customTextarea = card.querySelector('.custom-response-textarea');
    const useCustomCheckbox = card.querySelector('.use-custom-checkbox input[type="checkbox"]');
    
    if (aiTextarea) {
        aiTextarea.value = question.aiResponse || '';
        aiTextarea.disabled = question.isLoadingResponse;
        aiTextarea.placeholder = question.isLoadingResponse ? 'Generating AI response...' : 'AI response will appear here';
    }
    
    if (aiLoading) {
        aiLoading.style.display = question.isLoadingResponse ? 'flex' : 'none';
    }
    
    // Show "Don't like" button only when AI response exists, is not loading, and custom section is hidden
    if (dontLikeButton) {
        dontLikeButton.style.display = (!question.isLoadingResponse && question.aiResponse && !question.showCustomResponse) ? 'flex' : 'none';
    }
    
    // Show/hide custom response section
    if (customSection) {
        customSection.style.display = question.showCustomResponse ? 'block' : 'none';
    }
    
    // Update custom response value and checkbox
    if (customTextarea) {
        customTextarea.value = question.customResponse || '';
    }
    
    if (useCustomCheckbox) {
        useCustomCheckbox.checked = question.useCustomResponse;
    }
}

/**
 * Parse questions from email body using OpenAI (with fallback to pattern matching)
 * @param {Object} email - Email object with body content
 * @returns {Promise<Array>} Array of parsed questions
 */
async function parseAndDisplayQuestions(email) {
    const questionsList = document.getElementById('clarification-questions-list');
    const questionBox = document.getElementById('clarification-question-text');
    
    // Show loading state
    if (questionsList) {
        questionsList.innerHTML = '<div class="loading-indicator"><div class="spinner-small"></div><span>Parsing questions with AI...</span></div>';
        questionsList.classList.remove('hidden');
    }
    if (questionBox) questionBox.classList.add('hidden');
    
    let parsedQuestions = [];
    
    if (email.body?.content) {
        try {
            // Try OpenAI parsing first
            console.log('Attempting OpenAI question parsing...');
            parsedQuestions = await OpenAIService.parseQuestions(email.body.content, email.subject || '');
            console.log(`OpenAI parsed ${parsedQuestions.length} questions`);
        } catch (openaiError) {
            console.warn('OpenAI parsing failed, falling back to pattern matching:', openaiError);
            // Fallback to pattern matching
            try {
                parsedQuestions = Helpers.parseClarificationQuestions(email.body.content);
                console.log(`Pattern matching parsed ${parsedQuestions.length} questions`);
            } catch (patternError) {
                console.error('Both parsing methods failed:', patternError);
            }
        }
        
        // Initialize questions with empty responses
        AppState.questions = parsedQuestions.map((q, index) => ({
            id: `q${index + 1}`,
            question: q.question,
            category: q.category || 'General Questions',
            section_number: q.section_number || null,
            aiResponse: '',
            customResponse: '',
            useCustomResponse: false,
            showCustomResponse: false, // Only show custom response section when user doesn't like AI response
            isExpanded: false, // Start collapsed - user clicks to expand
            isLoadingResponse: false,
            hasError: false
        }));
        
        if (AppState.questions.length > 0) {
            // Render question cards
            renderQuestionCards(AppState.questions, email);
            
            // Generate AI responses for each question (async, don't await)
            generateAIResponsesForQuestions(AppState.questions, email).catch(err => {
                console.error('Error generating AI responses:', err);
            });
        } else {
            // No questions parsed - fallback to showing full body text
            if (questionsList) questionsList.classList.add('hidden');
            if (questionBox) {
                questionBox.classList.remove('hidden');
                const bodyText = Helpers.stripHtml(email.body.content);
                const truncatedBody = bodyText.length > 500 ? bodyText.substring(0, 500) + '...' : bodyText;
                questionBox.textContent = truncatedBody;
            }
        }
    } else {
        // No email body available
        if (questionsList) questionsList.classList.add('hidden');
        if (questionBox) {
            questionBox.classList.remove('hidden');
            questionBox.textContent = 'Email body not available';
        }
    }
}

/**
 * Show Clarification mode when user clicks on a clarification email
 */
async function showClarificationMode(context) {
    console.log('=== Showing Clarification mode ===');
    
    try {
        showMode('clarification-mode');
        AppState.currentMode = 'clarification';
        
        // Restore header title
        const headerTitle = document.getElementById('header-title');
        if (headerTitle) {
            headerTitle.textContent = '';
            headerTitle.removeAttribute('title');
        }
        
        const email = context.email;
        const originalRfq = context.originalRfq; // May be present if opened from sent RFQ
        
        // Store email context for button handlers
        AppState.emailContext = {
            email: email,
            originalRfq: originalRfq
        };
        
        if (!email) {
            console.error('No email data in context');
            // Still show the mode but with a message
            const emailInfoBox = document.getElementById('clarification-email-info');
            if (emailInfoBox) {
                emailInfoBox.innerHTML = '<p class="error-text">Could not load email details</p>';
            }
            return;
        }
        
        // Display email info
        const emailInfoBox = document.getElementById('clarification-email-info');
        if (emailInfoBox) {
            const fromAddress = email.from?.emailAddress?.address || 'Unknown sender';
            const fromName = email.from?.emailAddress?.name || fromAddress;
            const dateStr = email.receivedDateTime ? 
                new Date(email.receivedDateTime).toLocaleString() : 'Unknown date';
            
            let html = `
                <div class="email-subject">${Helpers.escapeHtml(email.subject || 'No subject')}</div>
                <div class="email-from">From: ${Helpers.escapeHtml(fromName)} &lt;${Helpers.escapeHtml(fromAddress)}&gt;</div>
                <div class="email-date">Received: ${dateStr}</div>
            `;
            
            // If we have the original RFQ context, show it
            if (originalRfq) {
                html += `
                    <div class="original-rfq-info" style="margin-top: 10px; padding-top: 10px; border-top: 1px solid #ddd; font-size: 12px; color: #666;">
                        <div><strong>In reply to your RFQ:</strong></div>
                        <div>${Helpers.escapeHtml(originalRfq.subject || 'Unknown subject')}</div>
                    </div>
                `;
            }
            
            emailInfoBox.innerHTML = html;
        }
        
        // Parse and display questions - this will extract questions and generate AI responses for each
        await parseAndDisplayQuestions(email);
        
    } catch (error) {
        console.error('Error in showClarificationMode:', error);
        Helpers.showError('Error displaying clarification: ' + error.message);
    }
}

/**
 * Process a clarification email via the proper backend API flow:
 * 1. Ensure email is classified (get backend email_id)
 * 2. Call /api/emails/process to get sub_classification and suggested_response
 * 3. If requires_engineering, show forward button; otherwise show suggested response
 */
async function processClarificationEmail(email) {
    const loadingEl = document.getElementById('suggested-answer-loading');
    const contentEl = document.getElementById('suggested-answer-content');
    const textareaEl = document.getElementById('clarification-response-text');
    const engineerBtnContainer = document.getElementById('forward-engineer-container');
    
    if (loadingEl) loadingEl.classList.remove('hidden');
    if (contentEl) contentEl.classList.add('hidden');
    if (engineerBtnContainer) engineerBtnContainer.classList.add('hidden');
    
    try {
        // Step 1: Ensure email is classified and get backend email_id
        console.log('Step 1: Ensuring email is classified...');
        const classifyResult = await ensureEmailClassified(email, 'clarification_request');
        const backendEmailId = classifyResult.emailId;
        console.log('Backend email_id:', backendEmailId);
        
        // Step 2: Call /api/emails/process to get sub_classification and suggested_response
        console.log('Step 2: Calling /api/emails/process...');
        let processResult;
        try {
            processResult = await ApiClient.processEmail(backendEmailId, 'clarification_request');
            console.log('Process result:', processResult);
            
            // Store clarification_id for future use (e.g., regenerating response)
            if (processResult.clarification_id) {
                storeClarificationId(email.id, processResult.clarification_id);
            }
        } catch (processError) {
            console.warn('/api/emails/process failed:', processError.message);
            // Fall back to using OpenAI directly to generate response
            try {
                console.log('Attempting to generate response using OpenAI directly...');
                
                // Check if OpenAI API key is configured
                if (!Config.OPENAI_API_KEY) {
                    throw new Error('OpenAI API key is not configured. Please set OPENAI_API_KEY in Vercel environment variables or localStorage.');
                }
                
                const questionText = email.body?.content ? Helpers.stripHtml(email.body.content).substring(0, 1000) : 'No question text available';
                const emailContext = {
                    subject: email.subject || '',
                    body: email.body?.content || '',
                    rfqContext: ''
                };
                const aiResponse = await OpenAIService.generateResponse(questionText, emailContext);
                processResult = {
                    suggested_response: aiResponse,
                    requires_engineering: false
                };
                console.log('OpenAI fallback response generated successfully');
            } catch (openaiError) {
                console.error('OpenAI fallback also failed:', openaiError);
                // Show helpful error message
                const errorMsg = openaiError.message && openaiError.message.includes('API key') 
                    ? 'OpenAI API key not configured. Please set OPENAI_API_KEY in Vercel environment variables.'
                    : `OpenAI error: ${openaiError.message || 'Unknown error'}`;
                console.error(errorMsg);
                // Final fallback to template response with error note
                processResult = {
                    suggested_response: generateFallbackResponse(email, email.body?.content || '') + 
                        '\n\n[Note: AI response generation failed. Please edit this response manually.]',
                    requires_engineering: false
                };
            }
        }
        
        // Step 3: Handle based on sub_classification
        // Always hide loading and show content, even if there's an error
        if (loadingEl) loadingEl.classList.add('hidden');
        if (contentEl) contentEl.classList.remove('hidden');
        
        if (processResult.requires_engineering) {
            // Show "Forward to Engineering" UI
            console.log('Clarification requires engineering review');
            if (engineerBtnContainer) {
                engineerBtnContainer.classList.remove('hidden');
            }
            if (textareaEl) {
                textareaEl.value = processResult.suggested_response || 
                    'This clarification requires engineering review. Please forward to your engineering team.';
                textareaEl.disabled = true;
            }
            // Store email info for forwarding
            AppState.currentClarification = {
                email: email,
                backendEmailId: backendEmailId,
                clarificationId: processResult.clarification_id,
                question: processResult.question
            };
        } else {
            // Show suggested response for procurement clarification
            console.log('Procurement clarification - showing suggested response');
            if (textareaEl) {
                textareaEl.value = processResult.suggested_response || 
                    generateFallbackResponse(email, email.body?.content || '');
                textareaEl.disabled = false;
            }
            // Store email info for sending reply
            AppState.currentClarification = {
                email: email,
                backendEmailId: backendEmailId,
                clarificationId: processResult.clarification_id,
                question: processResult.question
            };
        }
        
    } catch (error) {
        console.error('Error processing clarification email:', error);
        // Always hide loading and show content on error
        if (loadingEl) loadingEl.classList.add('hidden');
        if (contentEl) contentEl.classList.remove('hidden');
        
        // Try OpenAI fallback if backend completely fails
        if (textareaEl) {
            try {
                console.log('Attempting OpenAI fallback after error...');
                
                // Check if OpenAI API key is configured
                if (!Config.OPENAI_API_KEY) {
                    throw new Error('OpenAI API key is not configured. Please set OPENAI_API_KEY in Vercel environment variables.');
                }
                
                const questionText = email.body?.content ? Helpers.stripHtml(email.body.content).substring(0, 1000) : 'No question text available';
                const emailContext = {
                    subject: email.subject || '',
                    body: email.body?.content || '',
                    rfqContext: ''
                };
                const aiResponse = await OpenAIService.generateResponse(questionText, emailContext);
                textareaEl.value = aiResponse;
                console.log('OpenAI fallback succeeded');
            } catch (openaiError) {
                console.error('OpenAI fallback also failed:', openaiError);
                // Show helpful error message
                const errorMsg = openaiError.message && openaiError.message.includes('API key')
                    ? 'OpenAI API key not configured. Please set OPENAI_API_KEY in Vercel environment variables.'
                    : `Error: ${openaiError.message || error.message || 'Unknown error'}`;
                textareaEl.value = `[Error generating response: ${errorMsg}]\n\nPlease provide a custom response below:`;
                textareaEl.placeholder = 'Type your response here...';
                Helpers.showError(`Failed to generate AI response: ${errorMsg}`);
            }
        }
    }
}

/**
 * Load suggested response for a clarification email (legacy - now calls processClarificationEmail)
 * @deprecated Use processClarificationEmail instead
 */
async function loadSuggestedResponse(email) {
    return processClarificationEmail(email);
}

/**
 * Generate a fallback response template when API is unavailable
 */
function generateFallbackResponse(email, bodyText) {
    const senderName = email.from?.emailAddress?.name || 'Supplier';
    const subject = email.subject || 'your inquiry';
    
    return `Dear ${senderName},

Thank you for your email regarding "${subject}".

We have reviewed your questions and will respond with the requested information shortly.

[Please add your specific response here]

Best regards,
Procurement Team`;
}

/**
 * Show Quote mode when user clicks on a quote email
 */
async function showQuoteMode(context) {
    console.log('=== Showing Quote mode ===');
    
    try {
        showMode('quote-mode');
        AppState.currentMode = 'quote';
        
        // Restore header title
        const headerTitle = document.getElementById('header-title');
        if (headerTitle) {
            headerTitle.textContent = '';
            headerTitle.removeAttribute('title');
        }
        
        const email = context.email;
        const originalRfq = context.originalRfq; // May be present if opened from sent RFQ
        
        if (!email) {
            console.error('No email data in context');
            const emailInfoBox = document.getElementById('quote-email-info');
            if (emailInfoBox) {
                emailInfoBox.innerHTML = '<p class="error-text">Could not load email details</p>';
            }
            return;
        }
        
        // Display email info
        const emailInfoBox = document.getElementById('quote-email-info');
        if (emailInfoBox) {
            const fromAddress = email.from?.emailAddress?.address || 'Unknown sender';
            const fromName = email.from?.emailAddress?.name || fromAddress;
            const dateStr = email.receivedDateTime ? 
                new Date(email.receivedDateTime).toLocaleString() : 'Unknown date';
            
            let html = `
                <div class="email-subject">${Helpers.escapeHtml(email.subject || 'No subject')}</div>
                <div class="email-from">From: ${Helpers.escapeHtml(fromName)} &lt;${Helpers.escapeHtml(fromAddress)}&gt;</div>
                <div class="email-date">Received: ${dateStr}</div>
            `;
            
            // If we have the original RFQ context, show it
            if (originalRfq) {
                html += `
                    <div class="original-rfq-info" style="margin-top: 10px; padding-top: 10px; border-top: 1px solid #ddd; font-size: 12px; color: #666;">
                        <div><strong>In reply to your RFQ:</strong></div>
                        <div>${Helpers.escapeHtml(originalRfq.subject || 'Unknown subject')}</div>
                    </div>
                `;
            }
            
            emailInfoBox.innerHTML = html;
        }
        
        // Parse and display quote data (don't await - let it load in background)
        loadParsedQuoteData(email).catch(err => {
            console.error('Error loading quote data:', err);
        });
        
    } catch (error) {
        console.error('Error in showQuoteMode:', error);
        Helpers.showError('Error displaying quote: ' + error.message);
    }
}

/**
 * Load and parse quote data from email using proper API flow:
 * 1. Ensure email is classified (get backend email_id)
 * 2. Call /api/emails/process to confirm quote is ready
 * 3. Call /api/emails/extract-quote to get structured data
 */
async function loadParsedQuoteData(email) {
    const loadingEl = document.getElementById('quote-loading');
    const dataEl = document.getElementById('parsed-quote-data');
    
    if (loadingEl) loadingEl.classList.remove('hidden');
    if (dataEl) dataEl.classList.add('hidden');
    
    const supplierEmail = email.from?.emailAddress?.address || '';
    const supplierName = email.from?.emailAddress?.name || supplierEmail;
    const bodyContent = email.body?.content || '';
    const bodyText = Helpers.stripHtml(bodyContent);
    
    // Helper to safely set element text
    const setField = (id, value) => {
        const el = document.getElementById(id);
        if (el) el.textContent = value || '-';
    };
    
    // Try to extract basic info from email body directly (fallback)
    const extractFromBody = () => {
        const details = {
            supplier_name: supplierName,
            unit_price: null,
            total_price: null,
            lead_time: null,
            validity: null,
            payment_terms: null,
            notes: null
        };
        
        // Try to extract price from body
        const priceMatch = bodyText.match(/(?:unit\s*price|price)[:\s]*\$?([\d,]+\.?\d*)/i);
        if (priceMatch) details.unit_price = priceMatch[1];
        
        const totalMatch = bodyText.match(/(?:total\s*price|total)[:\s]*\$?([\d,]+\.?\d*)/i);
        if (totalMatch) details.total_price = totalMatch[1];
        
        const leadTimeMatch = bodyText.match(/(?:delivery|lead\s*time)[:\s]*([^\n\r]+)/i);
        if (leadTimeMatch) details.lead_time = leadTimeMatch[1].trim();
        
        const validityMatch = bodyText.match(/(?:validity|valid\s*for|quote\s*valid)[:\s]*([^\n\r]+)/i);
        if (validityMatch) details.validity = validityMatch[1].trim();
        
        const termsMatch = bodyText.match(/(?:payment\s*terms|terms)[:\s]*([^\n\r]+)/i);
        if (termsMatch) details.payment_terms = termsMatch[1].trim();
        
        return details;
    };
    
    try {
        // Extract RFQ ID from subject
        const rfqId = EmailOperations.extractRfqId ? 
            EmailOperations.extractRfqId(email.subject) : 
            (email.subject?.match(/MAT-\d+/)?.[0] || null);
        
        let details = {};
        
        // Try API with proper flow
        try {
            // Step 1: Ensure email is classified and get backend email_id
            console.log('Step 1: Ensuring quote email is classified...');
            const classifyResult = await ensureEmailClassified(email, 'quote');
            const backendEmailId = classifyResult.emailId;
            console.log('Backend email_id:', backendEmailId);
            
            // Step 2: Call /api/emails/process to confirm quote is ready
            console.log('Step 2: Calling /api/emails/process for quote...');
            try {
                const processResult = await ApiClient.processEmail(backendEmailId, 'quote');
                console.log('Process result:', processResult);
            } catch (processError) {
                // Process might fail but we can still try to extract
                console.warn('/api/emails/process failed (continuing anyway):', processError.message);
            }
            
            // Step 3: Call /api/emails/extract-quote to get structured data
            console.log('Step 3: Calling /api/emails/extract-quote...');
            const result = await ApiClient.extractQuote(
                backendEmailId,
                rfqId,
                supplierEmail,
                bodyContent
            );
            details = result.extracted_details || result || {};
            console.log('Quote extracted via API:', details);
            
            // Store quote info for later use
            AppState.currentQuote = {
                email: email,
                backendEmailId: backendEmailId,
                quoteId: result.quote_id,
                rfqId: rfqId,
                details: details
            };
        } catch (apiError) {
            console.warn('API quote extraction failed, using fallback:', apiError.message);
            details = extractFromBody();
            console.log('Quote extracted from body:', details);
        }
        
        if (loadingEl) loadingEl.classList.add('hidden');
        if (dataEl) dataEl.classList.remove('hidden');
        
        // Populate quote fields
        setField('quote-supplier', details.supplier_name || supplierName);
        setField('quote-price', details.unit_price ? `$${details.unit_price}` : null);
        setField('quote-total-price', details.total_price ? `$${details.total_price}` : null);
        setField('quote-leadtime', details.lead_time || details.delivery_time);
        setField('quote-validity', details.validity || details.quote_validity);
        setField('quote-terms', details.payment_terms);
        setField('quote-notes', details.notes || details.additional_notes);
        
    } catch (error) {
        console.error('Error parsing quote:', error);
        if (loadingEl) loadingEl.classList.add('hidden');
        if (dataEl) dataEl.classList.remove('hidden');
        
        // Try fallback extraction
        const fallbackDetails = extractFromBody();
        setField('quote-supplier', fallbackDetails.supplier_name || supplierName);
        setField('quote-price', fallbackDetails.unit_price ? `$${fallbackDetails.unit_price}` : null);
        setField('quote-total-price', fallbackDetails.total_price ? `$${fallbackDetails.total_price}` : null);
        setField('quote-leadtime', fallbackDetails.lead_time);
        setField('quote-validity', fallbackDetails.validity);
        setField('quote-terms', fallbackDetails.payment_terms);
        setField('quote-notes', 'Note: Quote details extracted from email body (API unavailable)');
    }
}

// ==================== MODE ACTION HANDLERS ====================

/**
 * Handle sending all RFQ drafts from Draft mode
 * CRITICAL: Complete ALL work (sending, folder moves, auto-replies) for non-current drafts
 * BEFORE touching the current draft which will close the add-in
 */
async function handleSendAllDraftsFromDraftMode() {
    if (!AuthService.isSignedIn()) {
        Helpers.showError('Please sign in to send emails');
        return;
    }
    
    const sendBtn = document.getElementById('send-all-drafts-btn');
    const draftListSection = document.getElementById('draft-list-section');
    const draftActionsSection = document.getElementById('draft-actions-section');
    const progressTracker = document.getElementById('rfq-progress-tracker');
    
    // Progress tracker elements
    const sentRfqCount = document.getElementById('sent-rfq-count');
    const sentRfqProgress = document.getElementById('sent-rfq-progress');
    const autoRepliesScheduledCount = document.getElementById('auto-replies-scheduled-count');
    const autoRepliesScheduledProgress = document.getElementById('auto-replies-scheduled-progress');
    const repliesReceivedCount = document.getElementById('replies-received-count');
    const repliesReceivedProgress = document.getElementById('replies-received-progress');
    const repliesSortedCount = document.getElementById('replies-sorted-count');
    const repliesSortedProgress = document.getElementById('replies-sorted-progress');
    
    try {
        // Hide draft list, show progress tracker
        if (draftListSection) draftListSection.classList.add('hidden');
        if (draftActionsSection) draftActionsSection.classList.add('hidden');
        if (progressTracker) progressTracker.classList.remove('hidden');
        if (sendBtn) sendBtn.disabled = true;
        
        // Initialize progress bars
        if (sentRfqCount) sentRfqCount.textContent = '0';
        if (sentRfqProgress) sentRfqProgress.style.width = '0%';
        if (autoRepliesScheduledCount) autoRepliesScheduledCount.textContent = '0';
        if (autoRepliesScheduledProgress) autoRepliesScheduledProgress.style.width = '0%';
        if (repliesReceivedCount) repliesReceivedCount.textContent = '0 / 0';
        if (repliesReceivedProgress) repliesReceivedProgress.style.width = '0%';
        if (repliesSortedCount) repliesSortedCount.textContent = '0 / 0';
        if (repliesSortedProgress) repliesSortedProgress.style.width = '0%';
        
        // Get all RFQ drafts
        let draftsResponse;
        try {
            draftsResponse = await AuthService.graphRequest(
                `/me/mailFolders/Drafts/messages?$filter=startswith(subject,'RFQ for')&$select=id,subject,toRecipients,body&$top=50`
            );
        } catch (filterError) {
            // Try without filter
            draftsResponse = await AuthService.graphRequest(
                `/me/mailFolders/Drafts/messages?$select=id,subject,toRecipients,body&$top=50`
            );
            if (draftsResponse.value) {
                draftsResponse.value = draftsResponse.value.filter(d => 
                    d.subject && d.subject.toLowerCase().startsWith('rfq for')
                );
            }
        }
        
        const drafts = draftsResponse.value || [];
        
        if (drafts.length === 0) {
            Helpers.showError('No RFQ drafts found');
            if (draftListSection) draftListSection.classList.remove('hidden');
            if (draftActionsSection) draftActionsSection.classList.remove('hidden');
            if (progressTracker) progressTracker.classList.add('hidden');
            if (sendBtn) sendBtn.disabled = false;
            return;
        }
        
        console.log(`Found ${drafts.length} RFQ drafts to send`);
        
        // Get current email ID (the draft we might be viewing)
        const currentDraftId = Office.context.mailbox.item?.itemId;
        console.log('Currently viewing draft ID:', currentDraftId);
        
        // Separate current draft from others - we'll send it LAST
        const otherDrafts = drafts.filter(d => d.id !== currentDraftId);
        const currentDraft = drafts.find(d => d.id === currentDraftId);
        
        const totalDrafts = drafts.length;
        
        // Track material codes for all drafts being sent
        const sentMaterialCodes = new Set();
        
        // Persist initial state
        persistState({
            sendingInProgress: true,
            totalDrafts: totalDrafts,
            sentCount: 0,
            autoRepliesScheduled: 0,
            materialCodes: []
        });
        
        let sentCount = 0;
        let autoRepliesScheduled = 0;
        
        // Update progress function
        const updateProgress = () => {
            // Update sent RFQs
            if (sentRfqCount) sentRfqCount.textContent = sentCount.toString();
            if (sentRfqProgress) sentRfqProgress.style.width = `${(sentCount / totalDrafts) * 100}%`;
            
            // Update auto-replies scheduled
            if (autoRepliesScheduledCount) autoRepliesScheduledCount.textContent = autoRepliesScheduled.toString();
            if (autoRepliesScheduledProgress) autoRepliesScheduledProgress.style.width = `${(autoRepliesScheduled / totalDrafts) * 100}%`;
            
            // Update replies received (will be updated by monitoring)
            if (repliesReceivedCount) repliesReceivedCount.textContent = `0 / ${sentCount}`;
            if (repliesReceivedProgress) repliesReceivedProgress.style.width = '0%';
            
            // Update replies sorted
            if (repliesSortedCount) repliesSortedCount.textContent = '0 / 0';
            if (repliesSortedProgress) repliesSortedProgress.style.width = '0%';
        };
        
        // STEP 1: Send ALL OTHER drafts first (not the current one)
        for (const draft of otherDrafts) {
            try {
                const recipient = draft.toRecipients?.[0]?.emailAddress?.address || 'unknown';
                console.log(`Sending to ${recipient}... (${sentCount + 1}/${totalDrafts})`);
                
                // Extract material code from draft subject
                const materialMatch = (draft.subject || '').match(/MAT-\d+/i);
                if (materialMatch) {
                    sentMaterialCodes.add(materialMatch[0].toUpperCase());
                    console.log(`  Material code: ${materialMatch[0].toUpperCase()}`);
                }
                
                // Send the draft and get the sent email details
                const sendResult = await sendDraftEmailWithFullWorkflow(draft);
                sentCount++;
                
                if (sendResult.autoReplyScheduled) {
                    autoRepliesScheduled++;
                }
                
                // Update progress bars
                updateProgress();
                
                // Update persisted state after each successful send
                persistState({ 
                    sentCount, 
                    autoRepliesScheduled,
                    materialCodes: Array.from(sentMaterialCodes)
                });
                
                console.log(`✓ Sent ${sentCount}/${totalDrafts}: ${draft.subject}`);
                
            } catch (error) {
                console.error(`✗ Failed to send draft to ${draft.toRecipients?.[0]?.emailAddress?.address}:`, error);
            }
        }
        
        // STEP 2: If we sent all non-current drafts and there's no current draft, we're done
        if (!currentDraft) {
            const materialCodesArray = Array.from(sentMaterialCodes);
            console.log(`Tracking replies for material codes: ${materialCodesArray.join(', ')}`);
            
            persistState({ 
                sendingInProgress: false, 
                lastSendResult: 'success',
                sentCount,
                autoRepliesScheduled,
                materialCodes: materialCodesArray
            });
            updateProgress();
            
            // Start monitoring for replies with material codes
            startReplyMonitoring(sentCount, updateProgress, materialCodesArray);
            
            Helpers.showSuccess(`Sent ${sentCount} RFQ(s) successfully! ${autoRepliesScheduled} auto-replies scheduled.`);
            return;
        }
        
        // Extract material code from current draft
        const currentMaterialMatch = (currentDraft.subject || '').match(/MAT-\d+/i);
        if (currentMaterialMatch) {
            sentMaterialCodes.add(currentMaterialMatch[0].toUpperCase());
            console.log(`  Current draft material code: ${currentMaterialMatch[0].toUpperCase()}`);
        }
        
        // STEP 3: Send the CURRENT draft last
        // After this, the add-in WILL close because we're viewing this draft
        console.log('Sending final draft... Panel will close shortly.');
        
        const materialCodesArray = Array.from(sentMaterialCodes);
        
        // Mark state as complete BEFORE sending current draft (because we won't get a chance after)
        persistState({ 
            sendingInProgress: false, 
            lastSendResult: 'success',
            sentCount: sentCount + 1, // Include the one we're about to send
            autoRepliesScheduled: autoRepliesScheduled + 1, // Assume it will work
            materialCodes: materialCodesArray,
            showSuccessOnReopen: true
        });
        
        // Update progress to show final state
        sentCount++;
        autoRepliesScheduled++;
        updateProgress();
        
        // Small delay so user sees the message
        await new Promise(resolve => setTimeout(resolve, 500));
        
        // Send the current draft - this will trigger add-in close
        try {
            await sendDraftEmailWithFullWorkflow(currentDraft);
            console.log('✓ Sent current draft successfully');
        } catch (error) {
            console.error('Error sending current draft:', error);
            // Try to update state even though add-in might close
            persistState({ lastSendResult: 'partial' });
        }
        
    } catch (error) {
        console.error('Error in send all drafts:', error);
        Helpers.showError('Error sending drafts: ' + error.message);
        persistState({ sendingInProgress: false, lastSendResult: 'error', errorMessage: error.message });
        
        // Restore UI on error
        if (draftListSection) draftListSection.classList.remove('hidden');
        if (draftActionsSection) draftActionsSection.classList.remove('hidden');
        if (progressTracker) progressTracker.classList.add('hidden');
        if (sendBtn) sendBtn.disabled = false;
    }
}

/**
 * Start monitoring for replies and update progress bars
 */
function startReplyMonitoring(totalSent, updateProgressCallback, materialCodes = []) {
    if (!AuthService.isSignedIn()) return;
    
    // Get material codes from state if not provided (for legacy/fallback)
    if (materialCodes.length === 0) {
        const state = getPersistedState();
        materialCodes = state.materialCodes || [];
        if (materialCodes.length === 0 && totalSent > 0) {
            console.warn('No material codes provided or in state - counting all folders (may include old replies)');
        }
    }
    
    if (materialCodes.length > 0) {
        console.log(`Reply monitoring: Filtering for material codes: ${materialCodes.join(', ')}`);
    }
    
    const repliesReceivedCount = document.getElementById('replies-received-count');
    const repliesReceivedProgress = document.getElementById('replies-received-progress');
    const repliesSortedCount = document.getElementById('replies-sorted-count');
    const repliesSortedProgress = document.getElementById('replies-sorted-progress');
    
    // Helper to check if email is an undeliverable/bounceback
    const isUndeliverable = (email, bodyPreview = '') => {
        const subject = (email.subject || '').toLowerCase();
        const from = (email.from?.emailAddress?.address || '').toLowerCase();
        const fromName = (email.from?.emailAddress?.name || '').toLowerCase();
        const body = (bodyPreview || '').toLowerCase();
        
        // Subject/from checks (existing)
        if (subject.includes('undeliverable') || 
            subject.includes('delivery failure') ||
            subject.includes('delivery has failed') ||
            subject.includes('mail delivery failed') ||
            from.includes('postmaster') ||
            from.includes('mailer-daemon') ||
            (from.includes('noreply') && subject.includes('failed')) ||
            fromName.includes('postmaster') ||
            fromName.includes('mailer-daemon')) {
            return true;
        }
        
        // NEW: Body content checks for bounceback patterns
        if (body.includes('message undeliverable') ||
            body.includes('delivery has failed') ||
            body.includes('returned mail') ||
            body.includes('mail delivery subsystem') ||
            body.includes('delivery status notification') ||
            body.includes('this is an automatically generated delivery status notification') ||
            body.includes('delivery to the following recipient failed') ||
            body.includes('could not be delivered')) {
            return true;
        }
        
        return false;
    };
    
    // Helper to check if email is a real supplier reply
    const isSupplierReply = (email, bodyPreview = '') => {
        const subject = (email.subject || '').toLowerCase();
        
        // Must contain RFQ in subject
        if (!subject.includes('rfq')) {
            return false;
        }
        
        // Must not be undeliverable
        if (isUndeliverable(email, bodyPreview)) {
            return false;
        }
        
        // NEW: Must have actual content (not just bounceback)
        const body = (bodyPreview || '').trim();
        if (body.length < 50) {
            return false; // Too short to be a real reply
        }
        
        // Check body doesn't contain bounceback patterns
        const bouncePatterns = [
            'delivery failed',
            'undeliverable',
            'returned mail',
            'mail delivery subsystem',
            'delivery status notification',
            'could not be delivered',
            'permanent failure',
            'temporary failure'
        ];
        if (bouncePatterns.some(pattern => body.includes(pattern))) {
            return false;
        }
        
        return true;
    };
    
    let checkInterval = setInterval(async () => {
        try {
            let repliesReceived = 0;
            let repliesSorted = 0;
            const allReplyIds = new Set(); // Track unique reply IDs
            
            // Step 1: Get all mail folders and find material code folders first
            let quoteFolders = [];
            let clarificationFolders = [];
            
            try {
                const allFolders = await AuthService.graphRequest(
                    `/me/mailFolders?$select=id,displayName,parentFolderId&$top=500`
                );
                
                if (allFolders.value) {
                    // First find material code folders (MAT-XXXXX pattern)
                    // Only include folders for material codes that were sent in the current batch
                    const materialFolders = [];
                    for (const folder of allFolders.value) {
                        if (/^MAT-\d+$/i.test(folder.displayName)) {
                            const folderCode = folder.displayName.toUpperCase();
                            // Only include if materialCodes is empty (legacy/fallback) or if this material code was tracked
                            if (materialCodes.length === 0 || materialCodes.includes(folderCode)) {
                                materialFolders.push(folder);
                            }
                        }
                    }
                    
                    // Then find Quotes/Clarification folders within each material folder
                    for (const materialFolder of materialFolders) {
                        try {
                            const materialSubfolders = await AuthService.graphRequest(
                                `/me/mailFolders/${materialFolder.id}/childFolders?$select=id,displayName&$top=20`
                            );
                            
                            if (materialSubfolders.value) {
                                for (const subfolder of materialSubfolders.value) {
                                    const name = (subfolder.displayName || '').toLowerCase();
                                    if (name.includes('quote') && !name.includes('sent')) {
                                        quoteFolders.push(subfolder.id);
                                    }
                                    if (name.includes('clarification') && !name.includes('awaiting')) {
                                        clarificationFolders.push(subfolder.id);
                                    }
                                }
                            }
                        } catch (e) {
                            // Ignore errors fetching subfolders for a specific material folder
                            console.warn(`Error fetching subfolders for ${materialFolder.displayName}:`, e);
                        }
                    }
                    
                    // Also check for top-level folders as fallback (in case structure is different)
                    for (const folder of allFolders.value) {
                        const folderName = (folder.displayName || '').toLowerCase();
                        if (folderName.includes('quote') && !folderName.includes('sent') && !quoteFolders.includes(folder.id)) {
                            quoteFolders.push(folder.id);
                        }
                        if (folderName.includes('clarification') && !folderName.includes('awaiting') && !clarificationFolders.includes(folder.id)) {
                            clarificationFolders.push(folder.id);
                        }
                    }
                }
            } catch (e) {
                console.warn('Error getting folders:', e);
            }
            
            // Step 2: Count emails in sorted folders (Quotes and Clarifications)
            for (const folderId of quoteFolders) {
                try {
                    const folderEmails = await AuthService.graphRequest(
                        `/me/mailFolders/${folderId}/messages?$select=id,subject,from,bodyPreview&$top=100`
                    );
                    if (folderEmails.value) {
                        for (const email of folderEmails.value) {
                            const bodyPreview = email.bodyPreview || '';
                            if (isSupplierReply(email, bodyPreview) && !allReplyIds.has(email.id)) {
                                allReplyIds.add(email.id);
                                repliesReceived++;
                                repliesSorted++;
                            }
                        }
                    }
                } catch (e) {
                    console.warn(`Error counting emails in quote folder ${folderId}:`, e);
                }
            }
            
            for (const folderId of clarificationFolders) {
                try {
                    const folderEmails = await AuthService.graphRequest(
                        `/me/mailFolders/${folderId}/messages?$select=id,subject,from,bodyPreview&$top=100`
                    );
                    if (folderEmails.value) {
                        for (const email of folderEmails.value) {
                            const bodyPreview = email.bodyPreview || '';
                            if (isSupplierReply(email, bodyPreview) && !allReplyIds.has(email.id)) {
                                allReplyIds.add(email.id);
                                repliesReceived++;
                                repliesSorted++;
                            }
                        }
                    }
                } catch (e) {
                    console.warn(`Error counting emails in clarification folder ${folderId}:`, e);
                }
            }
            
            // Step 3: Count RFQ-related emails in inbox (not yet sorted)
            try {
                const inboxReplies = await AuthService.graphRequest(
                    `/me/mailFolders/inbox/messages?$filter=contains(subject,'RFQ')&$top=100&$select=id,subject,from,bodyPreview&$orderby=receivedDateTime desc`
                );
                
                if (inboxReplies.value) {
                    for (const email of inboxReplies.value) {
                        const bodyPreview = email.bodyPreview || '';
                        if (isSupplierReply(email, bodyPreview) && !allReplyIds.has(email.id)) {
                            allReplyIds.add(email.id);
                            repliesReceived++;
                            // Not sorted yet, so don't increment repliesSorted
                        }
                    }
                }
            } catch (e) {
                console.warn('Error getting inbox replies:', e);
            }
            
            console.log(`Reply monitoring: ${repliesReceived} received, ${repliesSorted} sorted out of ${totalSent} sent`);
            
            // Update UI
            if (repliesReceivedCount) repliesReceivedCount.textContent = `${repliesReceived} / ${totalSent}`;
            if (repliesReceivedProgress && totalSent > 0) {
                repliesReceivedProgress.style.width = `${Math.min(100, (repliesReceived / totalSent) * 100)}%`;
            }
            
            if (repliesSortedCount) repliesSortedCount.textContent = `${repliesSorted} / ${repliesReceived}`;
            if (repliesSortedProgress && repliesReceived > 0) {
                repliesSortedProgress.style.width = `${Math.min(100, (repliesSorted / repliesReceived) * 100)}%`;
            } else if (repliesSortedProgress && repliesReceived === 0) {
                repliesSortedProgress.style.width = '0%';
            }
            
            // Stop monitoring if all replies are sorted
            if (repliesReceived >= totalSent && repliesSorted >= repliesReceived) {
                console.log('All replies received and sorted - stopping monitoring');
                clearInterval(checkInterval);
            }
            
        } catch (error) {
            console.warn('Error monitoring replies:', error);
        }
    }, 5000); // Check every 5 seconds
    
    // Stop after 10 minutes (give more time for replies)
    setTimeout(() => {
        clearInterval(checkInterval);
        console.log('Reply monitoring stopped after timeout');
    }, 10 * 60 * 1000);
}

/**
 * Send a single draft email with COMPLETE workflow:
 * Uses EmailOperations.sendEmail which has robust retry logic for:
 * 1. Send the email
 * 2. Find it in Sent Items (with multiple strategies)
 * 3. Move to correct folder
 * 4. Apply category
 * 5. Schedule auto-reply
 */
async function sendDraftEmailWithFullWorkflow(draft) {
    const subject = draft.subject || '';
    const recipient = draft.toRecipients?.[0]?.emailAddress?.address || '';
    const materialMatch = subject.match(/MAT-\d+/i);
    const materialCode = materialMatch ? materialMatch[0] : null;
    const body = draft.body?.content || '';
    
    console.log(`=== Sending draft: ${subject} to ${recipient} ===`);
    console.log(`Material code: ${materialCode}`);
    
    // Delete the draft first (to avoid duplicates), then send via sendEmail
    // Actually, we need to send the draft itself via Graph API, not create a new email
    
    // Step 1: Send the draft via Graph API
    console.log(`Sending draft ${draft.id}...`);
    await AuthService.graphRequest(`/me/messages/${draft.id}/send`, {
        method: 'POST'
    });
    console.log(`✓ Draft sent successfully`);
    
    // Step 2: Wait and find the sent email using the proven getSentEmails method
    let sentEmail = null;
    let internetMessageId = null;
    let movedEmailId = null;
    
    const maxRetries = 5;
    const initialDelay = 2000;
    
    for (let attempt = 0; attempt < maxRetries && !sentEmail; attempt++) {
        const delay = initialDelay + (attempt * 1000);
        console.log(`Waiting ${delay}ms for email to appear in Sent Items (attempt ${attempt + 1}/${maxRetries})...`);
        await new Promise(resolve => setTimeout(resolve, delay));
        
        // Strategy 1: Find by subject using getSentEmails
        const sentEmailsBySubject = await EmailOperations.getSentEmails({
            subject: subject,
            top: 10
        });
        
        if (sentEmailsBySubject.length > 0) {
            // Filter by recipient to ensure we get the right one
            for (const email of sentEmailsBySubject) {
                const toRecipients = email.toRecipients || [];
                const matchesRecipient = toRecipients.some(r => 
                    r.emailAddress?.address?.toLowerCase() === recipient.toLowerCase()
                );
                
                if (matchesRecipient) {
                    sentEmail = email;
                    console.log(`✓ Found sent email on attempt ${attempt + 1} by subject + recipient match`);
                    break;
                }
            }
            
            // If no recipient match, use the most recent one
            if (!sentEmail && sentEmailsBySubject.length > 0) {
                sentEmail = sentEmailsBySubject[0];
                console.log(`✓ Found sent email on attempt ${attempt + 1} by subject match`);
            }
        }
        
        // Strategy 2: If not found, try getting most recent emails
        if (!sentEmail) {
            const recentEmails = await EmailOperations.getSentEmails({ top: 20 });
            for (const email of recentEmails) {
                const toRecipients = email.toRecipients || [];
                const matchesRecipient = toRecipients.some(r => 
                    r.emailAddress?.address?.toLowerCase() === recipient.toLowerCase()
                );
                const matchesSubject = email.subject === subject;
                
                if (matchesRecipient && matchesSubject) {
                    sentEmail = email;
                    console.log(`✓ Found sent email on attempt ${attempt + 1} by recent emails search`);
                    break;
                }
            }
        }
    }
    
    if (!sentEmail) {
        console.error(`✗ Could not find sent email after ${maxRetries} attempts`);
        return { success: true, sentEmailId: null, internetMessageId: null, autoReplyScheduled: false };
    }
    
    internetMessageId = sentEmail.internetMessageId;
    console.log(`Sent email ID: ${sentEmail.id}, internetMessageId: ${internetMessageId}`);
    
    // Step 3: Move to correct folder if we have material code
    if (materialCode) {
        try {
            // Ensure folder structure exists
            console.log(`Initializing folder structure for ${materialCode}...`);
            await FolderManagement.initializeMaterialFolders(materialCode);
            console.log(`✓ Folder structure ready`);
            
            // Move to Sent RFQs folder
            const folderPath = `${materialCode}/${Config.FOLDERS.SENT_RFQS}`;
            console.log(`Moving email to ${folderPath}...`);
            const moveResult = await FolderManagement.moveEmailToFolder(sentEmail.id, folderPath);
            movedEmailId = moveResult?.id || sentEmail.id;
            console.log(`✓ Moved email to ${folderPath}`);
            
            // Wait for move to complete
            await new Promise(resolve => setTimeout(resolve, 1000));
            
            // Apply category to the moved email
            console.log(`Applying SENT RFQ category...`);
            await EmailOperations.applyCategoryToEmail(movedEmailId, 'SENT RFQ', 'Preset6');
            console.log(`✓ Category applied`);
        } catch (folderError) {
            console.error('Error with folder/category operations:', folderError);
            // Continue even if folder operations fail
        }
    }
    
    // Step 4: Schedule auto-reply for demo/testing
    let autoReplyScheduled = false;
    if (internetMessageId) {
        try {
            const userEmail = Office.context.mailbox.userProfile?.emailAddress;
            if (userEmail) {
                const quantityMatch = subject.match(/(\d+)\s*pcs/i);
                const quantity = quantityMatch ? parseInt(quantityMatch[1]) : 100;
                
                console.log(`Scheduling auto-reply to ${userEmail}...`);
                await ApiClient.scheduleAutoReply({
                    toEmail: userEmail,
                    subject: subject,
                    internetMessageId: internetMessageId,
                    material: materialCode || 'Unknown Material',
                    replyType: 'random',
                    delaySeconds: 5,
                    quantity: quantity
                });
                
                console.log(`✓ Auto-reply scheduled (will arrive in ~5 seconds)`);
                autoReplyScheduled = true;
            } else {
                console.warn('Could not get user email for auto-reply scheduling');
            }
        } catch (autoReplyError) {
            console.error('Error scheduling auto-reply:', autoReplyError);
            // Continue even if auto-reply fails
        }
    } else {
        console.warn('No internetMessageId available - cannot schedule auto-reply');
    }
    
    console.log(`=== Completed workflow for: ${subject} ===\n`);
    
    return {
        success: true,
        sentEmailId: movedEmailId || sentEmail.id,
        internetMessageId,
        autoReplyScheduled
    };
}

/**
 * Send a single draft email (simple version for backwards compatibility)
 */
async function sendDraftEmail(draft) {
    // Extract material code from subject
    const materialMatch = draft.subject?.match(/MAT-\d+/i);
    const materialCode = materialMatch ? materialMatch[0] : null;
    
    // Send the draft
    await AuthService.graphRequest(`/me/messages/${draft.id}/send`, {
        method: 'POST'
    });
    
    console.log(`Sent draft: ${draft.subject}`);
    
    // If we have a material code, try to move to the correct folder
    // Note: The email is already sent, so we'd need to find it in Sent Items
    // This is handled by EmailOperations.sendEmail in the regular flow
}

/**
 * Handle sending clarification to engineer
 */
async function handleSendToEngineer() {
    if (!AppState.emailContext?.email) {
        Helpers.showError('No email context');
        return;
    }
    
    if (!AuthService.isSignedIn()) {
        Helpers.showError('Please sign in to send emails');
        return;
    }
    
    const email = AppState.emailContext.email;
    
    // CRITICAL: Check if email is from Microsoft Outlook and delete immediately
    if (EmailOperations.isFromMicrosoftOutlook(email)) {
        try {
            Helpers.showLoading('Deleting Microsoft Outlook email...');
            if (email.id) {
                await EmailOperations.deleteEmail(email.id);
            }
            Helpers.showSuccess('Microsoft Outlook email deleted');
            Helpers.hideLoading();
            return;
        } catch (deleteError) {
            Helpers.showError('Failed to delete Microsoft Outlook email: ' + deleteError.message);
            Helpers.hideLoading();
            return;
        }
    }
    
    try {
        Helpers.showLoading('Forwarding to engineering...');
        const engineeringEmail = Config.getSetting(Config.STORAGE_KEYS.ENGINEERING_EMAIL, 'engineering@company.com');
        
        // Format HTML comment with questions
        let comment = `
            <p>Please review the following technical clarification request:</p>
            <hr>
            <p><strong>Original Email:</strong></p>
            <p><strong>From:</strong> ${email.from?.emailAddress?.address || 'Unknown'}</p>
            <p><strong>Subject:</strong> ${email.subject || 'No subject'}</p>
            <hr>
        `;
        
        // Add parsed questions if available
        if (AppState.questions && AppState.questions.length > 0) {
            comment += `<p><strong>Questions:</strong></p><ol>`;
            AppState.questions.forEach((q, index) => {
                comment += `<li><strong>${Helpers.escapeHtml(q.question)}</strong>`;
                if (q.category && q.category !== 'General Questions') {
                    comment += ` <em>(${Helpers.escapeHtml(q.category)})</em>`;
                }
                comment += `</li>`;
            });
            comment += `</ol><hr>`;
        }
        
        // Add original email body
        comment += `<p><strong>Original Email Body:</strong></p>${email.body?.content || 'No body content'}`;
        
        // Forward email directly (no compose window)
        await EmailOperations.forwardEmail(email.id, [engineeringEmail], comment);
        
        // Move original email to Awaiting Engineer folder after successful send
        const materialMatch = email.subject?.match(/MAT-\d+/i);
        if (materialMatch) {
            const folderPath = `${materialMatch[0]}/${Config.FOLDERS.AWAITING_ENGINEER}`;
            try {
                await FolderManagement.moveEmailToFolder(email.id, folderPath);
            } catch (e) {
                console.error('Could not move email to folder:', e);
                // Don't fail the operation - email was sent successfully
            }
        }
        
        Helpers.showSuccess('Email forwarded to engineering team');
        
        // Go back to workflow
        showRFQWorkflowMode();
        
    } catch (error) {
        Helpers.showError('Failed to forward: ' + error.message);
    } finally {
        Helpers.hideLoading();
    }
}

/**
 * Format clarification response with all questions and answers
 * @param {Array} questions - Array of question objects with responses
 * @param {Object} email - Email context
 * @returns {string} Formatted HTML response
 */
function formatClarificationResponse(questions, email) {
    const supplierName = email.from?.emailAddress?.name || 'Supplier';
    const supplierEmail = email.from?.emailAddress?.address || '';
    
    let html = `
        <p>Dear ${Helpers.escapeHtml(supplierName)},</p>
        <p>Thank you for your questions. Please find our responses below:</p>
        <br>
    `;
    
    questions.forEach((q, index) => {
        const questionNum = index + 1;
        const response = q.useCustomResponse && q.customResponse.trim() 
            ? q.customResponse.trim() 
            : (q.aiResponse || 'Response pending');
        
        html += `
            <div style="margin-bottom: 16px;">
                <p><strong>Q${questionNum}:</strong> ${Helpers.escapeHtml(q.question)}</p>
                <p><strong>A${questionNum}:</strong> ${Helpers.escapeHtml(response)}</p>
            </div>
        `;
    });
    
    html += `
        <br>
        <p>If you have any further questions, please don't hesitate to reach out.</p>
        <p>Best regards,<br>Procurement Team</p>
    `;
    
    return html;
}

/**
 * Handle replying to supplier with clarification response
 */
async function handleReplyToSupplier() {
    if (!AppState.emailContext?.email) {
        Helpers.showError('No email context');
        return;
    }
    
    if (!AuthService.isSignedIn()) {
        Helpers.showError('Please sign in to send emails');
        return;
    }
    
    const email = AppState.emailContext.email;
    
    // CRITICAL: Check if email is from Microsoft Outlook and delete immediately
    if (EmailOperations.isFromMicrosoftOutlook(email)) {
        try {
            Helpers.showLoading('Deleting Microsoft Outlook email...');
            if (email.id) {
                await EmailOperations.deleteEmail(email.id);
            }
            Helpers.showSuccess('Microsoft Outlook email deleted');
            Helpers.hideLoading();
            return;
        } catch (deleteError) {
            Helpers.showError('Failed to delete Microsoft Outlook email: ' + deleteError.message);
            Helpers.hideLoading();
            return;
        }
    }
    
    // Check if we have questions with responses
    if (!AppState.questions || AppState.questions.length === 0) {
        // Fallback to old textarea method (but still send directly)
        const responseText = document.getElementById('clarification-response-text')?.value;
        if (!responseText || responseText.trim().length === 0) {
            Helpers.showError('Please enter a response or wait for questions to be parsed');
            return;
        }
        
        try {
            Helpers.showLoading('Sending reply...');
            
            const email = AppState.emailContext.email;
            const htmlBody = EmailOperations.formatTextAsHtml(responseText);
            
            // Reply directly (no compose window)
            await EmailOperations.replyToEmail(email.id, htmlBody, false);
            
            // Move original email to Awaiting Clarification Response folder
            const materialMatch = email.subject?.match(/MAT-\d+/i);
            if (materialMatch) {
                const folderPath = `${materialMatch[0]}/${Config.FOLDERS.AWAITING_CLARIFICATION}`;
                try {
                    await FolderManagement.moveEmailToFolder(email.id, folderPath);
                } catch (e) {
                    console.error('Could not move email to folder:', e);
                }
            }
            
            Helpers.showSuccess('Reply sent to supplier');
            showRFQWorkflowMode();
        } catch (error) {
            Helpers.showError('Failed to send reply: ' + error.message);
        } finally {
            Helpers.hideLoading();
        }
        return;
    }
    
    // Validate that all questions have responses
    const questionsWithoutResponses = AppState.questions.filter(q => {
        const hasResponse = q.useCustomResponse 
            ? q.customResponse.trim().length > 0
            : (q.aiResponse && q.aiResponse.trim().length > 0);
        return !hasResponse;
    });
    
    if (questionsWithoutResponses.length > 0) {
        Helpers.showError(`Please provide responses for all questions. ${questionsWithoutResponses.length} question(s) still need responses.`);
        return;
    }
    
    try {
        Helpers.showLoading('Sending reply...');
        
        const email = AppState.emailContext.email;
        
        // Format response with all Q&A
        const htmlBody = formatClarificationResponse(AppState.questions, email);
        
        // Reply directly (no compose window)
        await EmailOperations.replyToEmail(email.id, htmlBody, false);
        
        // Move original email to Awaiting Clarification Response folder after successful send
        const materialMatch = email.subject?.match(/MAT-\d+/i);
        if (materialMatch) {
            const folderPath = `${materialMatch[0]}/${Config.FOLDERS.AWAITING_CLARIFICATION}`;
            try {
                await FolderManagement.moveEmailToFolder(email.id, folderPath);
            } catch (e) {
                console.error('Could not move email to folder:', e);
                // Don't fail the operation - email was sent successfully
            }
        }
        
        Helpers.showSuccess('Reply sent to supplier with all Q&A');
        
        // Go back to workflow
        showRFQWorkflowMode();
        
    } catch (error) {
        Helpers.showError('Failed to send reply: ' + error.message);
    } finally {
        Helpers.hideLoading();
    }
}

/**
 * Handle accepting a quote
 */
async function handleAcceptQuote() {
    if (!AppState.emailContext?.email) {
        Helpers.showError('No email context');
        return;
    }
    
    try {
        Helpers.showLoading('Processing quote acceptance...');
        
        // For now, just show a success message
        // In a real implementation, this would create a PO or update the system
        Helpers.showSuccess('Quote acceptance noted. Please create a Purchase Order in your ERP system.');
        
    } catch (error) {
        Helpers.showError('Failed to process: ' + error.message);
    } finally {
        Helpers.hideLoading();
    }
}

// ==================== INITIALIZATION ====================
// Global error handler to prevent crashes
window.onerror = function(message, source, lineno, colno, error) {
    console.error('Global error caught:', message, source, lineno, colno);
    console.error('Error object:', error);
    return false;
};

window.onunhandledrejection = function(event) {
    console.error('Unhandled promise rejection:', event.reason);
};

Office.onReady((info) => {
    console.log('Office.onReady fired, host:', info.host);
    if (info.host === Office.HostType.Outlook) {
        console.log('Office.js is ready in Outlook');
        initializeApp().catch(err => {
            console.error('initializeApp error:', err);
            // Try to show something
            const mainContent = document.getElementById('main-content');
            if (mainContent) mainContent.style.display = 'block';
        });
    } else {
        console.log('Running outside of Outlook - limited functionality');
        initializeApp().catch(err => {
            console.error('initializeApp error:', err);
        });
    }
});

async function initializeApp() {
    console.log('=== Initializing Procurement Workflow Add-in ===');
    
    try {
    // Load saved settings
    Config.loadSettings();
    
    // Fix any stuck overlays or modals
    const loadingOverlay = document.getElementById('loading-overlay');
    if (loadingOverlay && !loadingOverlay.classList.contains('hidden')) {
        console.log('Found stuck loading overlay, closing it...');
        loadingOverlay.classList.add('hidden');
    }
    
    // Close any stuck modals
    document.querySelectorAll('.modal').forEach(modal => {
        if (!modal.classList.contains('hidden')) {
            console.log('Found stuck modal, closing it...', modal.id);
            modal.classList.add('hidden');
        }
    });
    
        // Set up event listeners FIRST (so UI is responsive)
    setupEventListeners();
        setupModeEventListeners();
        
        // Initialize authentication - MUST await to ensure auth before context detection
        await initializeAuth();
    
        // Register for ItemChanged event
        registerItemChangedHandler();
        
        // Restore persisted state (shows success message if we were sending)
        try {
            restorePersistedState();
        } catch (e) {
            console.error('Error restoring state:', e);
        }
        
        // ALWAYS detect email context and render appropriate UI
        // Draft mode only shows when viewing a draft email
        // Normal/workflow mode shows for everything else
        try {
            console.log('Detecting email context...');
            const context = await detectEmailContext();
            console.log('Context detected:', context.type);
            await renderContextUI(context);
            console.log('Context UI rendered');
        } catch (contextError) {
            console.error('Error in context detection/rendering:', contextError);
            console.error('Stack:', contextError.stack);
            // ALWAYS show normal mode if there's any error
            showRFQWorkflowMode();
        }
        
        // Load initial data for workflow mode
        if (AppState.currentMode === 'rfq-workflow' || AppState.currentMode === 'normal' || !AppState.currentMode) {
            try {
    loadInitialData();
            } catch (dataError) {
                console.error('Error loading initial data:', dataError);
            }
        }
        
        console.log('=== Add-in initialization complete ===');
        
    } catch (fatalError) {
        console.error('FATAL ERROR during initialization:', fatalError);
        console.error('Stack:', fatalError.stack);
        // Ensure add-in shows something
        try {
            showRFQWorkflowMode();
        } catch (e) {
            // Last resort - show main content directly
            const mainContent = document.getElementById('main-content');
            if (mainContent) mainContent.style.display = 'block';
        }
    }
}

/**
 * Set up event listeners for mode-specific buttons
 */
function setupModeEventListeners() {
    // Back buttons
    document.getElementById('back-to-workflow-from-draft')?.addEventListener('click', () => {
        showRFQWorkflowMode();
        loadInitialData();
    });
    document.getElementById('back-to-workflow-from-clarification')?.addEventListener('click', () => {
        showRFQWorkflowMode();
        loadInitialData();
    });
    document.getElementById('back-to-workflow-from-quote')?.addEventListener('click', () => {
        showRFQWorkflowMode();
        loadInitialData();
    });
    
    // Draft mode buttons
    document.getElementById('send-all-drafts-btn')?.addEventListener('click', handleSendAllDraftsFromDraftMode);
    document.getElementById('view-draft-details-btn')?.addEventListener('click', showDraftDetailsModal);
    document.getElementById('close-draft-details-modal')?.addEventListener('click', closeDraftDetailsModal);
    document.getElementById('refresh-progress-btn')?.addEventListener('click', async () => {
        const state = getPersistedState();
        await loadRfqProgress(state);
        Helpers.showSuccess('Progress refreshed');
    });
    
    // Clarification mode buttons
    document.getElementById('send-to-engineer-btn')?.addEventListener('click', handleSendToEngineer);
    document.getElementById('reply-to-supplier-btn')?.addEventListener('click', handleReplyToSupplier);
    
    // Quote mode buttons
    document.getElementById('compare-quotes-btn')?.addEventListener('click', async () => {
        // Open quote comparison modal
        await openQuoteComparisonModal();
    });
    document.getElementById('accept-quote-btn')?.addEventListener('click', handleAcceptQuote);
    
    // Quote comparison modal event handlers
    document.getElementById('close-quote-comparison-modal')?.addEventListener('click', closeQuoteComparisonModal);
    document.getElementById('close-quote-comparison-modal-footer')?.addEventListener('click', closeQuoteComparisonModal);
    
    // Close modal on backdrop click
    document.getElementById('quote-comparison-modal')?.addEventListener('click', (e) => {
        if (e.target.id === 'quote-comparison-modal') {
            closeQuoteComparisonModal();
        }
    });
    
    // Close modal on Escape key
    document.addEventListener('keydown', (e) => {
        if (e.key === 'Escape') {
            const modal = document.getElementById('quote-comparison-modal');
            if (modal && !modal.classList.contains('hidden')) {
                closeQuoteComparisonModal();
            }
        }
    });
    
    // Sort dropdown
    document.getElementById('quote-sort-select')?.addEventListener('change', (e) => {
        modalQuotesState.sortBy = e.target.value;
        applyModalFiltersAndSort();
    });
    
    // Search input
    document.getElementById('quote-search-input')?.addEventListener('input', (e) => {
        modalQuotesState.filters.search = e.target.value;
        applyModalFiltersAndSort();
    });
    
    // Quick filter buttons
    document.getElementById('filter-best-price')?.addEventListener('click', (e) => {
        modalQuotesState.filters.bestPrice = !modalQuotesState.filters.bestPrice;
        e.target.classList.toggle('active', modalQuotesState.filters.bestPrice);
        applyModalFiltersAndSort();
    });
    
    document.getElementById('filter-fastest-delivery')?.addEventListener('click', (e) => {
        modalQuotesState.filters.fastestDelivery = !modalQuotesState.filters.fastestDelivery;
        e.target.classList.toggle('active', modalQuotesState.filters.fastestDelivery);
        applyModalFiltersAndSort();
    });
    
    document.getElementById('clear-filters')?.addEventListener('click', () => {
        modalQuotesState.filters = {
            search: '',
            bestPrice: false,
            fastestDelivery: false
        };
        document.getElementById('quote-search-input').value = '';
        document.getElementById('filter-best-price')?.classList.remove('active');
        document.getElementById('filter-fastest-delivery')?.classList.remove('active');
        applyModalFiltersAndSort();
    });
    
    // Export buttons
    document.getElementById('export-csv-btn')?.addEventListener('click', () => {
        exportQuotesToCSV(modalQuotesState.filteredQuotes);
    });
    
    document.getElementById('export-pdf-btn')?.addEventListener('click', () => {
        exportQuotesToPDF(modalQuotesState.filteredQuotes);
    });
    
    // Accept quote from modal
    document.getElementById('accept-quote-from-modal')?.addEventListener('click', () => {
        // Get the first quote (or selected quote if we add selection later)
        if (modalQuotesState.filteredQuotes.length > 0) {
            handleAcceptQuoteFromModal(modalQuotesState.filteredQuotes[0]);
        } else {
            Helpers.showError('No quote selected');
        }
    });
}

/**
 * Register handler for ItemChanged event
 * This is REQUIRED for the pinnable taskpane feature to work properly
 * When the user navigates to a different email while the taskpane is pinned,
 * this handler will be called to update the UI
 */
function registerItemChangedHandler() {
    try {
        if (Office.context.mailbox) {
            Office.context.mailbox.addHandlerAsync(
                Office.EventType.ItemChanged,
                onItemChanged,
                function(asyncResult) {
                    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                        console.log('ItemChanged handler registered successfully - pinning is enabled');
                    } else {
                        console.error('Failed to register ItemChanged handler:', asyncResult.error);
                    }
                }
            );
        }
    } catch (error) {
        console.error('Error registering ItemChanged handler:', error);
    }
}

/**
 * Handle ItemChanged event when user navigates to a different email
 * This is called when the taskpane is pinned and user selects a different email
 */
async function onItemChanged(eventArgs) {
    console.log('Item changed - user navigated to different email');
    
    // Re-detect context and update UI
    try {
        const context = await detectEmailContext();
        await renderContextUI(context);
    } catch (error) {
        console.error('Error handling item change:', error);
        showRFQWorkflowMode();
    }
    
    // If in workflow mode, update email info if email processing is visible
    if (AppState.currentMode === 'rfq-workflow') {
        const emailProcessingTab = document.getElementById('email-processing-tab');
        if (emailProcessingTab && !emailProcessingTab.classList.contains('hidden')) {
            updateCurrentEmailInfo();
        }
    }
}

// ==================== AUTHENTICATION ====================
async function initializeAuth() {
    try {
        const initialized = await AuthService.initialize();
        if (initialized) {
            updateAuthUI();
            // Start email monitoring if signed in
            if (AuthService.isSignedIn()) {
                EmailMonitor.startMonitoring();
            }
        }
    } catch (error) {
        console.error('Auth initialization failed:', error);
    }
}

function updateAuthUI() {
    const signInBtn = document.getElementById('sign-in-btn');
    const userName = document.getElementById('user-name');
    const accountDivider = document.getElementById('account-divider');
    const signOutBtn = document.getElementById('sign-out-btn');

    if (AuthService.isSignedIn()) {
        const user = AuthService.getUser();
        signInBtn?.classList.add('hidden');
        userName?.classList.remove('hidden');
        accountDivider?.classList.remove('hidden');
        signOutBtn?.classList.remove('hidden');
        if (userName && user) {
            const userDisplayName = user.name || user.email;
            userName.textContent = userDisplayName;
            userName.setAttribute('title', userDisplayName);
        }
    } else {
        signInBtn?.classList.remove('hidden');
        userName?.classList.add('hidden');
        accountDivider?.classList.add('hidden');
        signOutBtn?.classList.add('hidden');
    }
}

async function handleSignIn() {
    try {
        Helpers.showLoading('Signing in...');
        await AuthService.signIn();
        updateAuthUI();
        
        // Start email monitoring after sign-in
        EmailMonitor.startMonitoring();
        
        Helpers.showSuccess('Signed in successfully');
    } catch (error) {
        console.error('Sign in error:', error);
        Helpers.showError('Sign in failed: ' + error.message);
    } finally {
        Helpers.hideLoading();
    }
}

async function handleSignOut() {
    try {
        // Stop email monitoring
        EmailMonitor.stopMonitoring();
        
        await AuthService.signOut();
        updateAuthUI();
        FolderManagement.clearCache();
        Helpers.showSuccess('Signed out');
    } catch (error) {
        console.error('Sign out error:', error);
    }
}

// ==================== EVENT LISTENERS ====================
function setupEventListeners() {
    // Auth buttons
    document.getElementById('sign-in-btn')?.addEventListener('click', handleSignIn);
    document.getElementById('sign-out-btn')?.addEventListener('click', handleSignOut);

    // Navigation tabs
    // Tab navigation removed - using context-based mode switching only

    // Refresh button
    document.getElementById('refresh-btn')?.addEventListener('click', handleRefresh);

    // Settings button
    document.getElementById('settings-btn')?.addEventListener('click', openSettingsModal);
    document.getElementById('close-settings')?.addEventListener('click', closeSettingsModal);
    document.getElementById('save-settings')?.addEventListener('click', saveSettings);

    // PR modal - with retry mechanism
    // PR Button - Use event delegation (works even if element is added dynamically)
    console.log('Setting up PR button event delegation...');
    
    // Event delegation on document (always works)
    document.addEventListener('click', (e) => {
        const target = e.target.closest('#pr-step-title');
        if (target) {
            e.preventDefault();
            e.stopPropagation();
            console.log('PR step title clicked (via delegation)');
            openPRModal();
        }
    });
    
    // Also try direct listener as backup
    const attachPRButtonListener = () => {
        const prStepTitle = document.getElementById('pr-step-title');
        if (prStepTitle) {
            prStepTitle.addEventListener('click', (e) => {
                e.preventDefault();
                e.stopPropagation();
                console.log('PR step title clicked (via direct listener)');
                openPRModal();
            });
            console.log('PR step title direct event listener attached');
            return true;
        } else {
            console.warn('PR step title element not found for direct listener');
            return false;
        }
    };
    
    // Try to attach direct listener immediately
    attachPRButtonListener();
    
    // Retry direct listener after delay
    setTimeout(() => {
        attachPRButtonListener();
    }, 500);
    document.getElementById('close-pr-modal')?.addEventListener('click', closePRModal);
    document.getElementById('apply-pr-selection')?.addEventListener('click', applyPRSelection);
    
    // Close PR modal on ESC key
    document.addEventListener('keydown', (e) => {
        if (e.key === 'Escape') {
            const prModal = document.getElementById('pr-selection-modal');
            if (prModal && !prModal.classList.contains('hidden')) {
                closePRModal();
            }
        }
    });
    document.getElementById('pr-search')?.addEventListener('input', 
        Helpers.debounce(handlePRSearch, 300));

    // Select all suppliers
    document.getElementById('select-all-suppliers')?.addEventListener('change', handleSelectAllSuppliers);

    // Supplier modal
    document.getElementById('supplier-step-title')?.addEventListener('click', openSupplierModal);
    document.getElementById('close-supplier-modal')?.addEventListener('click', closeSupplierModal);
    document.getElementById('apply-supplier-selection')?.addEventListener('click', applySupplierSelection);
    document.getElementById('add-new-supplier-btn')?.addEventListener('click', handleAddNewSupplier);
    document.getElementById('supplier-search')?.addEventListener('input', handleSupplierSearch);

    // Generate RFQs step
    document.getElementById('generate-rfqs-step-title')?.addEventListener('click', handleGenerateRFQs);

    // Email processing
    document.getElementById('classify-email-btn')?.addEventListener('click', handleClassifyEmail);
    document.getElementById('extract-quote-btn')?.addEventListener('click', handleExtractQuote);
    document.getElementById('send-response-btn')?.addEventListener('click', handleSendClarificationResponse);
    document.getElementById('forward-to-engineering-btn')?.addEventListener('click', handleForwardToEngineering);
    document.getElementById('process-engineer-response-btn')?.addEventListener('click', handleProcessEngineerResponse);
    document.getElementById('create-engineer-draft-btn')?.addEventListener('click', handleCreateEngineerDraft);

    // Quote comparison
    document.getElementById('refresh-quotes-btn')?.addEventListener('click', () => {
        loadAllQuotesFromFolder();
    });

    // RFQ Preview Modal
    document.getElementById('close-rfq-preview')?.addEventListener('click', closeRFQPreviewModal);
    document.getElementById('create-draft-btn')?.addEventListener('click', handleCreateDraft);
    document.getElementById('finalize-send-btn')?.addEventListener('click', handleFinalizeSend);

    // Dismiss notifications
    document.getElementById('dismiss-error')?.addEventListener('click', () => {
        Helpers.hideElement(document.getElementById('error-banner'));
    });
    document.getElementById('dismiss-success')?.addEventListener('click', () => {
        Helpers.hideElement(document.getElementById('success-banner'));
    });
    
    // Pin reminder banner - Use event delegation (works even if element is added dynamically)
    console.log('Setting up banner close button event delegation...');
    
    // Event delegation on document (always works)
    document.addEventListener('click', (e) => {
        // Check if click is on the dismiss button or its icon
        const dismissBtn = e.target.closest('#dismiss-pin-reminder');
        if (dismissBtn) {
            e.preventDefault();
            e.stopPropagation();
            console.log('Dismiss pin reminder button clicked (via delegation)');
            dismissPinReminder();
        }
    });
    
    // Also try direct listener as backup
    const attachBannerCloseListener = () => {
        const dismissBtn = document.getElementById('dismiss-pin-reminder');
        if (dismissBtn) {
            dismissBtn.addEventListener('click', (e) => {
                e.preventDefault();
                e.stopPropagation();
                console.log('Dismiss pin reminder button clicked (via direct listener)');
                dismissPinReminder();
            });
            console.log('Pin reminder dismiss button direct event listener attached');
            return true;
        } else {
            console.warn('Dismiss pin reminder button not found for direct listener');
            return false;
        }
    };
    
    // Try to attach direct listener immediately
    attachBannerCloseListener();
    
    // Retry direct listener after delay
    setTimeout(() => {
        attachBannerCloseListener();
    }, 500);
    
    // Check if pin reminder was previously dismissed
    initPinReminder();
}

/**
 * Initialize pin reminder banner - show it if user hasn't dismissed it
 */
function initPinReminder() {
    const pinReminderDismissed = localStorage.getItem('procurement_pin_reminder_dismissed');
    const pinReminderBanner = document.getElementById('pin-reminder-banner');
    
    if (pinReminderDismissed === 'true') {
        pinReminderBanner?.classList.add('hidden');
    } else {
        pinReminderBanner?.classList.remove('hidden');
    }
}

/**
 * Dismiss the pin reminder and remember the choice
 */
function dismissPinReminder() {
    console.log('dismissPinReminder called');
    try {
        const pinReminderBanner = document.getElementById('pin-reminder-banner');
        if (pinReminderBanner) {
            pinReminderBanner.classList.add('hidden');
            localStorage.setItem('procurement_pin_reminder_dismissed', 'true');
            console.log('Pin reminder dismissed and saved');
        } else {
            console.warn('Pin reminder banner not found');
        }
    } catch (error) {
        console.error('Error dismissing pin reminder:', error);
    }
}

// ==================== TAB NAVIGATION ====================
// Tab navigation removed - using context-based mode switching only

// ==================== DATA LOADING ====================
async function loadInitialData() {
    // Check if we have persisted state with PRs
    const state = getPersistedState();
    
    if (state.prs && state.prs.length > 0) {
        // Restore PRs from state
        AppState.prs = state.prs;
        renderPRList(AppState.prs);
        
        // Don't restore selected PR on initial page load - let user select manually
        // This ensures no PR is pre-selected when the page first opens
        AppState.selectedPR = null;
        return;
    }
    
    // Otherwise auto-load PRs
    try {
        Helpers.showLoading('Loading open PRs...');
        await loadOpenPRs();
    } catch (error) {
        console.log('Could not auto-load PRs:', error.message);
        // Show placeholder instead of error - user can click refresh to load
        const prList = document.getElementById('pr-list');
        if (prList) {
            prList.innerHTML = '<p class="placeholder-text">Click the <strong>Refresh</strong> button ↻ to load open Purchase Requisitions</p>';
        }
    } finally {
        Helpers.hideLoading();
    }
}

async function loadOpenPRs() {
    try {
        AppState.prs = await ApiClient.getOpenPRs();
        renderPRList(AppState.prs);
    } catch (error) {
        console.error('Error loading PRs:', error);
        throw error;
    }
}

async function handleRefresh() {
    try {
        Helpers.showLoading('Refreshing data...');
        await loadOpenPRs();
        Helpers.showSuccess('Data refreshed successfully');
    } catch (error) {
        Helpers.showError('Failed to refresh: ' + error.message);
    } finally {
        Helpers.hideLoading();
    }
}

function updateContextUI() {
    // Check if we're viewing an email
    try {
        if (Office.context && Office.context.mailbox && Office.context.mailbox.item) {
            loadCurrentEmailInfo();
        }
    } catch (e) {
        console.log('Not in email context');
    }
}

// ==================== PR LIST RENDERING ====================
function renderPRList(prs) {
    const container = document.getElementById('pr-list');
    if (!container) return;
    
    Helpers.clearChildren(container);
    
    if (prs.length === 0) {
        container.innerHTML = '<p class="placeholder-text">No open PRs found</p>';
        return;
    }
    
    prs.forEach(pr => {
        const item = Helpers.createElement('div', {
            className: 'list-item',
            dataset: { prId: pr.pr_id },
            onClick: () => handlePRSelect(pr)
        }, `
            <div class="list-item-title">${Helpers.escapeHtml(pr.pr_id)}</div>
            <div class="list-item-subtitle">
                Material: ${Helpers.escapeHtml(pr.material || 'N/A')}
            </div>
            <div class="list-item-meta">
                Qty: ${pr.quantities || 'N/A'} ${pr.unit || ''} | 
                Created: ${Helpers.formatDate(pr.created_date)}
            </div>
        `);
        
        container.appendChild(item);
    });
}

function handlePRSearch(event) {
    const searchTerm = event.target.value;
    const filtered = Helpers.filterBySearch(
        AppState.prs, 
        searchTerm, 
        'pr_id', 'material', 'description'
    );
    renderPRList(filtered);
}

function openPRModal() {
    console.log('openPRModal called');
    try {
        const modal = document.getElementById('pr-selection-modal');
        if (!modal) {
            console.error('PR modal not found');
            return;
        }
        
        // Ensure any stuck loading overlays are closed
        Helpers.hideLoading();
        
        // Load PRs if not already loaded
        if (!AppState.prs || AppState.prs.length === 0) {
            loadOpenPRs().then(() => {
                if (AppState.prs && AppState.prs.length > 0) {
                    renderPRList(AppState.prs);
                }
            }).catch(error => {
                console.error('Error loading PRs:', error);
                Helpers.showError('Failed to load PRs: ' + error.message);
            });
        } else {
            renderPRList(AppState.prs);
        }
        
        // Restore selection if PR is already selected
        if (AppState.selectedPR) {
            restorePRSelection();
        }
        
        // Show modal
        modal.classList.remove('hidden');
        
        // Close modal when clicking outside (on the modal backdrop)
        const modalContent = modal.querySelector('.modal-content');
        if (modalContent) {
            // Remove any existing click handlers to avoid duplicates
            modalContent.onclick = null;
            modal.onclick = (e) => {
                // If click is on the modal backdrop (not on modal-content), close it
                if (e.target === modal) {
                    closePRModal();
                }
            };
        }
        
        // Focus search input
        const searchInput = document.getElementById('pr-search');
        if (searchInput) {
            setTimeout(() => searchInput.focus(), 100);
        }
    } catch (error) {
        console.error('Error in openPRModal:', error);
        Helpers.showError('Failed to open PR selection: ' + error.message);
    }
}

function closePRModal() {
    console.log('closePRModal called');
    try {
        const modal = document.getElementById('pr-selection-modal');
        if (modal) {
            modal.classList.add('hidden');
        }
        // Also ensure loading overlay is closed
        Helpers.hideLoading();
    } catch (error) {
        console.error('Error in closePRModal:', error);
    }
}

function applyPRSelection() {
    // Close modal
    closePRModal();
}

function restorePRSelection() {
    // Highlight selected PR in modal
    if (AppState.selectedPR) {
        document.querySelectorAll('#pr-list .list-item').forEach(item => {
            item.classList.remove('selected');
            if (item.dataset.prId === AppState.selectedPR.pr_id) {
                item.classList.add('selected');
            }
        });
        
        // Show selected PR details in modal
        updatePRDetailsInModal();
    }
}

function updatePRDetailsInModal() {
    const detailsContainer = document.getElementById('pr-details');
    const infoContainer = document.getElementById('selected-pr-info');
    
    if (AppState.selectedPR && detailsContainer) {
        const description = AppState.selectedPR.description || 'N/A';
        detailsContainer.innerHTML = `
            <p class="pr-description">${Helpers.escapeHtml(description)}</p>
        `;
        if (infoContainer) infoContainer.classList.remove('hidden');
    } else {
        if (infoContainer) infoContainer.classList.add('hidden');
    }
}

// ==================== PR SELECTION ====================
async function handlePRSelect(pr) {
    // Update UI selection state
    document.querySelectorAll('#pr-list .list-item').forEach(item => {
        item.classList.remove('selected');
        if (item.dataset.prId === pr.pr_id) {
            item.classList.add('selected');
        }
    });
    
    AppState.selectedPR = pr;
    
    // Clear RFQs and selected suppliers when a new PR is selected (allows generating RFQs for new PR)
    AppState.rfqs = [];
    AppState.selectedSuppliers = [];
    
    // Remove success message if it exists
    const successMsg = document.querySelector('.rfq-generated-message');
    if (successMsg) {
        successMsg.remove();
    }
    
    // Reset button state
    updateGenerateRFQsStep();
    
    // Persist the selection
    persistState({ 
        selectedPR: pr,
        prs: AppState.prs,
        currentStep: 'suppliers'
    });
    
    // Update PR details in modal
    updatePRDetailsInModal();
    
    // Enable supplier step and load suppliers
    Helpers.enableStep(document.getElementById('step-select-suppliers'));
    // Step 3 (Generate RFQs) will be enabled when suppliers are selected (handled in updateGenerateRFQsStep)
    
    try {
        Helpers.showLoading('Searching for matching suppliers...');
        const suppliers = await ApiClient.searchSuppliers(pr.pr_id, pr.material);
        AppState.suppliers = suppliers;
        renderSupplierList(suppliers);
    } catch (error) {
        Helpers.showError('Failed to load suppliers: ' + error.message);
    } finally {
        Helpers.hideLoading();
    }
}

// ==================== SUPPLIER LIST RENDERING ====================
function renderSupplierList(suppliers) {
    const container = document.getElementById('supplier-list');
    const countEl = document.getElementById('supplier-count');
    
    if (!container) return;
    
    Helpers.clearChildren(container);
    // Don't clear selectedSuppliers - preserve selection when re-rendering
    // AppState.selectedSuppliers = [];
    updateGenerateRFQsStep();
    
    if (countEl) {
        countEl.textContent = `${suppliers.length} supplier${suppliers.length !== 1 ? 's' : ''} found`;
    }
    
    if (suppliers.length === 0) {
        container.innerHTML = '<p class="placeholder-text">No matching suppliers found</p>';
        return;
    }
    
    suppliers.forEach(supplier => {
        const isSelected = AppState.selectedSuppliers.includes(supplier.supplier_id);
        
        const item = Helpers.createElement('div', {
            className: `list-item ${isSelected ? 'selected' : ''}`,
            dataset: { supplierId: supplier.supplier_id },
            onClick: () => handleSupplierSelect(supplier.supplier_id)
        }, `
            <div class="supplier-info">
                <div class="list-item-title">
                    ${Helpers.escapeHtml(supplier.name)}
                    <span class="match-score">${supplier.match_score}/10</span>
                </div>
                <div class="list-item-subtitle">
                    ${Helpers.escapeHtml(supplier.email)}
                </div>
                <div class="list-item-meta">
                    Contact: ${Helpers.escapeHtml(supplier.contact_person || 'N/A')} |
                    ${Helpers.escapeHtml(supplier.match_reason || '')}
                </div>
            </div>
        `);
        
        container.appendChild(item);
    });
    
    updateGenerateRFQsStep();
}

function handleSupplierSelect(supplierId) {
    const item = document.querySelector(`[data-supplier-id="${supplierId}"]`);
    if (!item) return;
    
    const isSelected = AppState.selectedSuppliers.includes(supplierId);
    
    if (isSelected) {
        // Deselect
        AppState.selectedSuppliers = AppState.selectedSuppliers.filter(id => id !== supplierId);
        item.classList.remove('selected');
    } else {
        // Select
        AppState.selectedSuppliers.push(supplierId);
        item.classList.add('selected');
    }
    
    // Update select all checkbox
    const allItems = document.querySelectorAll('#supplier-list .list-item');
    const selectAllCheckbox = document.getElementById('select-all-suppliers');
    if (selectAllCheckbox) {
        selectAllCheckbox.checked = AppState.selectedSuppliers.length === allItems.length && allItems.length > 0;
    }
    
    updateSelectedSuppliersCount();
    updateGenerateRFQsStep();
}

// Keep old function for backward compatibility (if called from elsewhere)
function handleSupplierCheckboxChange(checkbox) {
    const supplierId = checkbox.dataset.supplierId;
    handleSupplierSelect(supplierId);
}

function handleSelectAllSuppliers(event) {
    const isChecked = event.target.checked;
    const items = document.querySelectorAll('#supplier-list .list-item');
    
    AppState.selectedSuppliers = [];
    
    items.forEach(item => {
        const supplierId = item.dataset.supplierId;
        if (isChecked) {
            AppState.selectedSuppliers.push(supplierId);
            item.classList.add('selected');
        } else {
            item.classList.remove('selected');
        }
    });
    
    updateSelectedSuppliersCount();
    updateGenerateRFQsStep();
}

function updateSelectedSuppliersCount() {
    const count = AppState.selectedSuppliers.length;
    const summaryEl = document.getElementById('selected-suppliers-count');
    const modalCountEl = document.getElementById('selected-count-display');
    
    if (summaryEl) {
        summaryEl.textContent = `${count} supplier${count !== 1 ? 's' : ''} selected`;
    }
    
    if (modalCountEl) {
        modalCountEl.textContent = `${count} selected`;
    }
}

function openSupplierModal() {
    const modal = document.getElementById('supplier-selection-modal');
    if (!modal) return;
    
    // Clear any previous selections when opening the modal
    // This ensures a clean state when the modal is first opened
    if (AppState.selectedSuppliers.length > 0 && !AppState.selectedPR) {
        AppState.selectedSuppliers = [];
    }
    
    // Render suppliers in modal if we have them
    if (AppState.suppliers && AppState.suppliers.length > 0) {
        renderSupplierList(AppState.suppliers);
    }
    
    // Restore checkbox states
    restoreSupplierCheckboxStates();
    
    // Update counts
    updateSelectedSuppliersCount();
    
    // Show modal
    modal.classList.remove('hidden');
    
    // Focus search input
    const searchInput = document.getElementById('supplier-search');
    if (searchInput) {
        setTimeout(() => searchInput.focus(), 100);
    }
}

function closeSupplierModal() {
    const modal = document.getElementById('supplier-selection-modal');
    if (modal) {
        modal.classList.add('hidden');
    }
}

function applySupplierSelection() {
    // Update summary and close modal
    updateSelectedSuppliersCount();
    closeSupplierModal();
    updateGenerateRFQsStep();
}

function handleAddNewSupplier() {
    // Non-functional button - just show a message
    Helpers.showSuccess('Add New Supplier feature coming soon');
}

function handleSupplierSearch(event) {
    const searchTerm = event.target.value.toLowerCase();
    const suppliers = AppState.suppliers || [];
    
    if (!searchTerm) {
        renderSupplierList(suppliers);
        return;
    }
    
    const filtered = suppliers.filter(supplier => {
        const name = (supplier.name || '').toLowerCase();
        const email = (supplier.email || '').toLowerCase();
        const contact = (supplier.contact_person || '').toLowerCase();
        const reason = (supplier.match_reason || '').toLowerCase();
        
        return name.includes(searchTerm) || 
               email.includes(searchTerm) || 
               contact.includes(searchTerm) || 
               reason.includes(searchTerm);
    });
    
    renderSupplierList(filtered);
}

function restoreSupplierCheckboxStates() {
    // Restore selected states based on AppState.selectedSuppliers
    const items = document.querySelectorAll('#supplier-list .list-item');
    items.forEach(item => {
        const supplierId = item.dataset.supplierId;
        if (AppState.selectedSuppliers.includes(supplierId)) {
            item.classList.add('selected');
        } else {
            item.classList.remove('selected');
        }
    });
    
    // Update select all checkbox
    const selectAllCheckbox = document.getElementById('select-all-suppliers');
    if (selectAllCheckbox && items.length > 0) {
        selectAllCheckbox.checked = AppState.selectedSuppliers.length === items.length;
    }
}

function updateGenerateRFQsStep() {
    const step = document.getElementById('step-generate-rfqs');
    if (step) {
        // Enable step if suppliers are selected and RFQs haven't been generated yet
        const hasGeneratedRFQs = AppState.rfqs && AppState.rfqs.length > 0;
        
        if (!hasGeneratedRFQs && AppState.selectedSuppliers.length > 0) {
            Helpers.enableStep(step);
        } else if (AppState.selectedSuppliers.length === 0) {
            step.classList.add('disabled');
        }
    }
}

// ==================== RFQ GENERATION ====================
async function handleGenerateRFQs() {
    if (!AppState.selectedPR || AppState.selectedSuppliers.length === 0) {
        Helpers.showError('Please select a PR and at least one supplier');
        return;
    }
    
    try {
        Helpers.showLoading('Generating RFQs...');
        
        const rfqs = await ApiClient.generateRFQs(
            AppState.selectedPR.pr_id,
            AppState.selectedSuppliers
        );
        
        // Auto-save drafts if user is signed in
        if (AuthService.isSignedIn()) {
            Helpers.showLoading('Preparing attachments...');
            
            // Prepare PDF attachments for Graph API
            let graphApiAttachments = [];
            try {
                console.log('Starting attachment preparation...');
                graphApiAttachments = await AttachmentUtils.prepareGraphApiAttachments();
                console.log(`✓ Prepared ${graphApiAttachments.length} attachment(s) for drafts`);
                
                // Verify attachments have content
                for (const att of graphApiAttachments) {
                    if (!att.contentBytes || att.contentBytes.length === 0) {
                        console.error(`⚠ Attachment ${att.name} has no content!`);
                    } else {
                        console.log(`✓ Attachment ${att.name} has ${att.contentBytes.length} bytes`);
                    }
                }
                
                if (graphApiAttachments.length === 0) {
                    console.warn('⚠ No attachments prepared - drafts will be created without PDFs');
                    Helpers.showError('Warning: Could not load PDF attachments. Drafts will be created without them.');
                }
            } catch (attachmentError) {
                console.error('✗ CRITICAL: Failed to prepare attachments:', attachmentError);
                console.error('Error stack:', attachmentError.stack);
                Helpers.showError('Failed to load PDF attachments: ' + attachmentError.message);
                // Continue without attachments if preparation fails
            }
            
            Helpers.showLoading(`Saving ${rfqs.length} draft(s) with ${graphApiAttachments.length} attachment(s)...`);
            
            let successCount = 0;
            let failCount = 0;
            
            for (const rfq of rfqs) {
                try {
                    // Extract body content and convert to HTML
                    const bodyContent = EmailOperations.extractBodyContent(rfq.body);
                    let htmlBody = '';
                    if (bodyContent && typeof bodyContent === 'string' && bodyContent.trim().length > 0) {
                        if (bodyContent.trim().toLowerCase().startsWith('<') && 
                            (bodyContent.includes('</') || bodyContent.includes('/>'))) {
                            htmlBody = bodyContent;
                        } else {
                            htmlBody = EmailOperations.formatTextAsHtml(bodyContent);
                        }
                    } else {
                        htmlBody = '<div>&nbsp;</div>';
                    }
                    
                    console.log(`Creating draft for ${rfq.supplier_name} with ${graphApiAttachments.length} attachment(s)...`);
                    
                    // Save draft with attachments
                    const draft = await EmailOperations.saveDraft({
                        to: [rfq.supplier_email],
                        subject: rfq.subject || '',
                        body: htmlBody,
                        cc: [],
                        attachments: graphApiAttachments
                    });
                    
                    // Store draft ID in RFQ object
                    rfq.draftId = draft.id;
                    successCount++;
                    console.log(`✓ Draft created for ${rfq.supplier_name} (ID: ${draft.id})`);
                } catch (error) {
                    failCount++;
                    console.error(`✗ Failed to save draft for ${rfq.supplier_name}:`, error);
                    console.error('Error stack:', error.stack);
                    // Continue with other RFQs even if one fails
                }
            }
            
            console.log(`Draft creation complete: ${successCount} succeeded, ${failCount} failed`);
        }
        
        AppState.rfqs = rfqs;
        
        // Keep the Generate RFQs step enabled (highlighted) after completion
        // This matches the behavior of other completed steps (Select PR, Select Suppliers)
        const generateStep = document.getElementById('step-generate-rfqs');
        if (generateStep) {
            // Remove disabled class if it exists, to ensure it stays highlighted
            generateStep.classList.remove('disabled');
            // The step remains clickable and highlighted like other completed steps
        }
        
        // Enable the Review & Send RFQs step
        const reviewStep = document.getElementById('step-review-rfqs');
        if (reviewStep) {
            Helpers.enableStep(reviewStep);
        }
        
        // Clear the RFQ list container and show success message
        const container = document.getElementById('rfq-list');
        if (container) {
            container.innerHTML = `
                <div style="text-align: center; padding: 40px 20px; color: #107c10;">
                    <div style="font-size: 48px; margin-bottom: 16px;">✓</div>
                    <div style="font-size: 18px; font-weight: 600; margin-bottom: 8px;">RFQs Generated</div>
                    <div style="font-size: 14px; color: #605e5c; margin-bottom: 16px;">${rfqs.length} RFQ draft(s) have been created and saved to your Drafts folder.</div>
                    <div style="font-size: 13px; color: #0078d4; font-weight: 500;">Check your Drafts folder in Outlook to review and send them.</div>
                </div>
            `;
        }
        
        Helpers.showSuccess(`${rfqs.length} RFQ(s) generated successfully`);
    } catch (error) {
        Helpers.showError('Failed to generate RFQs: ' + error.message);
    } finally {
        Helpers.hideLoading();
    }
}

function renderRFQCards(rfqs) {
    const container = document.getElementById('rfq-list');
    if (!container) return;
    
    Helpers.clearChildren(container);
    
    rfqs.forEach((rfq, index) => {
        const card = Helpers.createElement('div', {
            className: 'rfq-card',
            dataset: { rfqId: rfq.rfq_id, index: index }
        }, `
            <div class="rfq-card-header">
                <h4>${Helpers.escapeHtml(rfq.supplier_name)}</h4>
                <span class="rfq-status ${rfq.status}">${rfq.status}</span>
            </div>
            <div class="rfq-card-body">
                <p><strong>To:</strong> ${Helpers.escapeHtml(rfq.supplier_email)}</p>
                <p><strong>Subject:</strong> ${Helpers.escapeHtml(rfq.subject)}</p>
                <p><strong>Attachments:</strong> ${rfq.attachments?.length || 0} file(s)</p>
            </div>
            <div class="rfq-card-actions">
                <button class="ms-Button ms-Button--small" onclick="previewRFQ(${index})">
                    <span class="ms-Button-label">Preview & Edit</span>
                </button>
                <button class="ms-Button ms-Button--small ms-Button--primary" onclick="createSingleDraft(${index})">
                    <span class="ms-Button-label">${rfq.draftId ? 'View Draft' : 'Create Draft'}</span>
                </button>
            </div>
        `);
        
        container.appendChild(card);
    });
}

// ==================== RFQ PREVIEW & EDITING ====================
let currentPreviewIndex = null;

function previewRFQ(index) {
    const rfq = AppState.rfqs[index];
    if (!rfq) return;
    
    currentPreviewIndex = index;
    
    // Populate preview modal
    document.getElementById('preview-to').value = rfq.supplier_email;
    document.getElementById('preview-subject').value = rfq.subject;
    document.getElementById('preview-body').value = EmailOperations.formatRfqBodyAsText(rfq.body);
    
    // Show attachments
    const attachmentsContainer = document.getElementById('preview-attachments');
    if (attachmentsContainer) {
        if (rfq.attachments && rfq.attachments.length > 0) {
            attachmentsContainer.innerHTML = rfq.attachments.map(att => 
                `<div class="attachment-item"><i class="ms-Icon ms-Icon--Attach"></i> ${Helpers.escapeHtml(att)}</div>`
            ).join('');
        } else {
            attachmentsContainer.innerHTML = '<span class="placeholder-text">No attachments</span>';
        }
    }
    
    // Show modal
    Helpers.showElement(document.getElementById('rfq-preview-modal'));
}

function closeRFQPreviewModal() {
    Helpers.hideElement(document.getElementById('rfq-preview-modal'));
    currentPreviewIndex = null;
}

async function handleCreateDraft() {
    if (currentPreviewIndex === null) return;
    
    const rfq = AppState.rfqs[currentPreviewIndex];
    const editedSubject = document.getElementById('preview-subject').value;
    const editedBody = document.getElementById('preview-body').value;
    
    // Convert plain text back to HTML for the email
    const htmlBody = EmailOperations.formatTextAsHtml(editedBody);
    
    try {
        Helpers.showLoading('Creating draft...');
        
        const result = await EmailOperations.createDraft(
            rfq.supplier_email,
            editedSubject,
            htmlBody,
            rfq.attachments?.map(att => ({ name: att, url: att }))
        );
        
        if (result.status === 'draft_saved_and_opened') {
            Helpers.showSuccess('Draft saved to Drafts folder and opened for editing');
        } else {
        Helpers.showSuccess('Draft created successfully');
        }
        closeRFQPreviewModal();
    } catch (error) {
        Helpers.showError('Failed to create draft: ' + error.message);
    } finally {
        Helpers.hideLoading();
    }
}

async function handleFinalizeSend() {
    if (currentPreviewIndex === null) return;
    
    const rfq = AppState.rfqs[currentPreviewIndex];
    const editedSubject = document.getElementById('preview-subject').value;
    const editedBody = document.getElementById('preview-body').value;
    
    // Convert plain text back to HTML for the email
    const htmlBody = EmailOperations.formatTextAsHtml(editedBody);
    
    try {
        Helpers.showLoading('Finalizing and sending...');
        
        // First finalize with backend
        await ApiClient.finalizeRFQ(rfq.rfq_id, editedSubject, editedBody, 'ready_to_send');
        
        // Create the email draft (user will send from Outlook)
        await EmailOperations.createDraft(
            rfq.supplier_email,
            editedSubject,
            htmlBody,
            rfq.attachments?.map(att => ({ name: att, url: att }))
        );
        
        // Update RFQ status in state
        AppState.rfqs[currentPreviewIndex].status = 'ready';
        
        // Create folder structure and move email
        const materialCode = Helpers.extractMaterialCode(AppState.selectedPR);
        await FolderManagement.initializeMaterialFolders(materialCode);
        
        Helpers.showSuccess('RFQ finalized - please send from Outlook');
        closeRFQPreviewModal();
        renderRFQCards(AppState.rfqs);
    } catch (error) {
        Helpers.showError('Failed to finalize: ' + error.message);
    } finally {
        Helpers.hideLoading();
    }
}

async function createSingleDraft(index) {
    const rfq = AppState.rfqs[index];
    if (!rfq) return;
    
    try {
        // If draft already exists, open it
        if (rfq.draftId && AuthService.isSignedIn()) {
            try {
                Helpers.showLoading('Opening draft...');
                await EmailOperations.openDraft(rfq.draftId);
                Helpers.showSuccess('Draft opened for ' + rfq.supplier_name);
                return;
            } catch (error) {
                console.error('Failed to open existing draft, creating new one:', error);
                // Fall through to create new draft
            }
        }
        
        // Create new draft if no existing draft or opening failed
        Helpers.showLoading('Preparing attachments...');
        
        // Prepare attachments for this draft
        let attachments = [];
        try {
            if (AuthService.isSignedIn()) {
                // Use Graph API format for signed-in users
                attachments = await AttachmentUtils.prepareGraphApiAttachments();
            } else {
                // Use Office.js format for non-signed-in users
                attachments = await AttachmentUtils.prepareOfficeJsAttachments();
            }
        } catch (attachmentError) {
            console.warn('Failed to prepare attachments, continuing without them:', attachmentError);
            // Continue without attachments if preparation fails
        }
        
        Helpers.showLoading('Creating draft...');
        
        // Extract body content (handles both string and object formats)
        const bodyContent = EmailOperations.extractBodyContent(rfq.body);
        
        // Convert to HTML if needed
        let htmlBody = '';
        if (bodyContent && typeof bodyContent === 'string' && bodyContent.trim().length > 0) {
            // Check if already HTML
            if (bodyContent.trim().toLowerCase().startsWith('<') && 
                (bodyContent.includes('</') || bodyContent.includes('/>'))) {
                htmlBody = bodyContent;
            } else {
                // Convert plain text to HTML
                htmlBody = EmailOperations.formatTextAsHtml(bodyContent);
            }
        } else {
            // Default to empty HTML div if no body
            htmlBody = '<div>&nbsp;</div>';
        }
        
        const result = await EmailOperations.createDraft(
            rfq.supplier_email,
            rfq.subject,
            htmlBody,
            attachments
        );
        
        // Store draft ID if it was saved
        if (result.draftId) {
            rfq.draftId = result.draftId;
            // Re-render to update button label
            renderRFQCards(AppState.rfqs);
        }
        
        if (result.status === 'draft_saved_and_opened') {
            Helpers.showSuccess('Draft saved to Drafts folder and opened for ' + rfq.supplier_name);
        } else {
        Helpers.showSuccess('Draft created for ' + rfq.supplier_name);
        }
    } catch (error) {
        Helpers.showError('Failed to create draft: ' + error.message);
    } finally {
        Helpers.hideLoading();
    }
}

async function handleCreateAllDrafts() {
    try {
        Helpers.showLoading('Creating all drafts...');
        
        for (const rfq of AppState.rfqs) {
            // Extract body content (handles both string and object formats)
            const bodyContent = EmailOperations.extractBodyContent(rfq.body);
            
            // Convert to HTML if needed
            let htmlBody = '';
            if (bodyContent && typeof bodyContent === 'string' && bodyContent.trim().length > 0) {
                // Check if already HTML
                if (bodyContent.trim().toLowerCase().startsWith('<') && 
                    (bodyContent.includes('</') || bodyContent.includes('/>'))) {
                    htmlBody = bodyContent;
                } else {
                    // Convert plain text to HTML
                    htmlBody = EmailOperations.formatTextAsHtml(bodyContent);
                }
            } else {
                // Default to empty HTML div if no body
                htmlBody = '<div>&nbsp;</div>';
            }
            
            await EmailOperations.createDraft(
                rfq.supplier_email,
                rfq.subject,
                htmlBody
            );
        }
        
        Helpers.showSuccess(`${AppState.rfqs.length} draft(s) created successfully`);
    } catch (error) {
        Helpers.showError('Failed to create drafts: ' + error.message);
    } finally {
        Helpers.hideLoading();
    }
}

// ==================== EMAIL PROCESSING ====================
async function loadCurrentEmailInfo() {
    try {
        const emailDetails = await EmailOperations.getCurrentEmailDetails();
        AppState.currentEmail = emailDetails;
        
        const container = document.getElementById('current-email-info');
        if (container && emailDetails) {
            container.innerHTML = `
                <div class="email-subject">${Helpers.escapeHtml(emailDetails.subject)}</div>
                <div class="email-from">From: ${Helpers.escapeHtml(emailDetails.from)}</div>
                <div class="email-date">${Helpers.formatDate(emailDetails.date, true)}</div>
            `;
        }
    } catch (error) {
        console.log('Could not load email info:', error);
        const container = document.getElementById('current-email-info');
        if (container) {
            container.innerHTML = '<p class="placeholder-text">Select an email to process</p>';
        }
    }
}

async function handleClassifyEmail() {
    if (!AppState.currentEmail) {
        Helpers.showError('No email selected');
        return;
    }
    
    // CRITICAL: Check if email is from Microsoft Outlook and delete immediately
    const email = AppState.currentEmail;
    if (EmailOperations.isFromMicrosoftOutlook(email)) {
        try {
            Helpers.showLoading('Deleting Microsoft Outlook email...');
            if (email.id) {
                await EmailOperations.deleteEmail(email.id);
            }
            Helpers.showSuccess('Microsoft Outlook email deleted');
            Helpers.hideLoading();
            // Clear current email state
            AppState.currentEmail = null;
            return;
        } catch (deleteError) {
            Helpers.showError('Failed to delete Microsoft Outlook email: ' + deleteError.message);
            Helpers.hideLoading();
            return;
        }
    }
    
    try {
        Helpers.showLoading('Classifying email...');
        
        const emailChain = await EmailOperations.getEmailChain();
        const rfqId = EmailOperations.extractRfqId(AppState.currentEmail.subject);
        
        const result = await ApiClient.classifyEmail(
            emailChain,
            {
                subject: AppState.currentEmail.subject,
                body: AppState.currentEmail.body,
                from_email: AppState.currentEmail.from,
                date: AppState.currentEmail.date?.toISOString() || new Date().toISOString(),
                in_reply_to: rfqId
            },
            rfqId
        );
        
        AppState.classification = result;
        
        // Display classification result
        const classificationCard = document.getElementById('classification-result');
        const badge = document.getElementById('classification-badge');
        const confidence = document.getElementById('classification-confidence');
        
        if (badge) {
            badge.textContent = Helpers.getClassificationDisplayName(result.classification);
            badge.className = `classification-badge ${result.classification}`;
        }
        
        if (confidence) {
            confidence.textContent = Helpers.formatConfidence(result.confidence);
        }
        
        Helpers.showElement(classificationCard);
        
        // Show processing section
        Helpers.showElement(document.getElementById('processing-section'));
        
        // Show appropriate processing card based on classification
        Helpers.hideElement(document.getElementById('quote-processing'));
        Helpers.hideElement(document.getElementById('clarification-processing'));
        Helpers.hideElement(document.getElementById('engineer-response-processing'));
        
        switch (result.classification) {
            case 'quote':
                Helpers.showElement(document.getElementById('quote-processing'));
                break;
            case 'clarification_request':
                Helpers.showElement(document.getElementById('clarification-processing'));
                break;
            case 'engineer_response':
                Helpers.showElement(document.getElementById('engineer-response-processing'));
                break;
        }
        
        Helpers.showSuccess('Email classified as: ' + result.classification);
    } catch (error) {
        Helpers.showError('Failed to classify email: ' + error.message);
    } finally {
        Helpers.hideLoading();
    }
}

async function handleExtractQuote() {
    if (!AppState.classification) {
        Helpers.showError('Please classify the email first');
        return;
    }
    
    try {
        Helpers.showLoading('Extracting quote data...');
        
        const rfqId = EmailOperations.extractRfqId(AppState.currentEmail.subject);
        
        const result = await ApiClient.extractQuote(
            AppState.classification.email_id,
            rfqId,
            null, // Supplier ID would come from classification
            AppState.currentEmail.body
        );
        
        // Display extracted quote details
        const container = document.getElementById('extracted-quote-details');
        if (container && result.extracted_details) {
            const details = result.extracted_details;
            container.innerHTML = `
                <h4>Extracted Quote Details:</h4>
                <div class="quote-field">
                    <label>Price:</label>
                    <span class="value price">${Helpers.formatCurrency(details.price, details.currency)}</span>
                </div>
                <div class="quote-field">
                    <label>Delivery Time:</label>
                    <span class="value">${Helpers.escapeHtml(details.delivery_time || 'N/A')}</span>
                </div>
                <div class="quote-field">
                    <label>Validity:</label>
                    <span class="value">${Helpers.escapeHtml(details.validity || 'N/A')}</span>
                </div>
                <div class="quote-field">
                    <label>Terms:</label>
                    <span class="value">${Helpers.escapeHtml(details.terms || 'N/A')}</span>
                </div>
                <p class="mt-10"><strong>Quote ID:</strong> ${result.quote_id}</p>
            `;
            Helpers.showElement(container);
        }
        
        Helpers.showSuccess('Quote extracted successfully');
    } catch (error) {
        Helpers.showError('Failed to extract quote: ' + error.message);
    } finally {
        Helpers.hideLoading();
    }
}

async function handleSendClarificationResponse() {
    const responseText = document.getElementById('suggested-response')?.value;
    
    if (!responseText) {
        Helpers.showError('Please enter a response');
        return;
    }
    
    try {
        Helpers.showLoading('Creating reply...');
        
        await EmailOperations.createReplyDraft(responseText);
        
        Helpers.showSuccess('Reply draft created');
    } catch (error) {
        Helpers.showError('Failed to create reply: ' + error.message);
    } finally {
        Helpers.hideLoading();
    }
}

async function handleForwardToEngineering() {
    if (!AppState.classification) {
        Helpers.showError('Please classify the email first');
        return;
    }
    
    try {
        Helpers.showLoading('Forwarding to engineering...');
        
        await ApiClient.forwardToEngineering(
            AppState.classification.email_id,
            AppState.processingResult?.clarification_id
        );
        
        // Create forward email
        const subject = `[Engineering Review] ${AppState.currentEmail.subject}`;
        const bodyText = `Please review the following technical clarification request:\n\n${AppState.currentEmail.body || ''}`;
        const htmlBody = EmailOperations.formatTextAsHtml(bodyText);
        
        await EmailOperations.createDraft(
            Config.ENGINEERING_EMAIL,
            subject,
            htmlBody
        );
        
        Helpers.showSuccess('Forwarded to engineering team');
    } catch (error) {
        Helpers.showError('Failed to forward: ' + error.message);
    } finally {
        Helpers.hideLoading();
    }
}

async function handleProcessEngineerResponse() {
    try {
        Helpers.showLoading('Processing engineer response...');
        
        const result = await ApiClient.processEngineerResponse(
            AppState.classification?.email_id,
            {
                body: AppState.currentEmail.body,
                from: AppState.currentEmail.from
            }
        );
        
        // Show draft response
        const draftContainer = document.getElementById('engineer-draft-response');
        const draftBody = document.getElementById('engineer-draft-body');
        
        if (draftBody && result.draft_response) {
            draftBody.value = result.draft_response.body;
            AppState.processingResult = result;
            Helpers.showElement(draftContainer);
        }
        
        Helpers.showSuccess('Draft response generated');
    } catch (error) {
        Helpers.showError('Failed to process response: ' + error.message);
    } finally {
        Helpers.hideLoading();
    }
}

async function handleCreateEngineerDraft() {
    const draftBody = document.getElementById('engineer-draft-body')?.value;
    
    if (!draftBody) {
        Helpers.showError('No draft content');
        return;
    }
    
    try {
        Helpers.showLoading('Creating reply...');
        
        // Get the original sender's email
        const toEmail = AppState.currentEmail?.from?.address || AppState.currentEmail?.from;
        if (!toEmail) {
            throw new Error('No sender email found');
        }
        
        const subject = `RE: ${AppState.currentEmail?.subject || 'Response'}`;
        const htmlBody = EmailOperations.formatTextAsHtml(draftBody);
        
        await EmailOperations.createDraft(toEmail, subject, htmlBody);
        
        Helpers.showSuccess('Reply draft created');
    } catch (error) {
        Helpers.showError('Failed to create reply: ' + error.message);
    } finally {
        Helpers.hideLoading();
    }
}

// ==================== QUOTE COMPARISON ====================
async function loadAvailableRFQs() {
    // In a real implementation, this would fetch from the backend
    // For now, use the RFQs we've generated in this session
    const select = document.getElementById('rfq-select');
    if (!select) return;
    
    // Clear existing options except the first
    while (select.options.length > 1) {
        select.remove(1);
    }
    
    // Add RFQs
    AppState.rfqs.forEach(rfq => {
        const option = document.createElement('option');
        option.value = rfq.rfq_id;
        option.textContent = `${rfq.rfq_id} - ${rfq.supplier_name}`;
        select.appendChild(option);
    });
}

/**
 * Show quote comparison view and automatically load all quotes
 */
async function showQuoteComparisonView() {
    // Hide all modes
    hideAllModes();
    
    // Show quote comparison section
    const mainContent = document.getElementById('main-content');
    const quoteComparisonTab = document.getElementById('quote-comparison-tab');
    
    if (quoteComparisonTab && mainContent) {
        // Hide other tab sections
        document.querySelectorAll('.tab-content').forEach(tab => {
            tab.classList.remove('active');
            tab.classList.add('hidden');
        });
        
        // Show quote comparison
        quoteComparisonTab.classList.remove('hidden');
        quoteComparisonTab.classList.add('active');
        mainContent.style.display = 'block';
        
        // Automatically load all quotes
        await loadAllQuotesFromFolder();
    }
}

/**
 * Get material code from current email's folder context
 * Checks if email is in a Quotes folder under a material folder (MAT-XXXXX/Quotes)
 */
async function getMaterialCodeFromEmailContext() {
    try {
        // Try to get current email from Office.js context first
        let currentEmailId = null;
        try {
            if (Office.context.mailbox && Office.context.mailbox.item && Office.context.mailbox.item.itemId) {
                currentEmailId = Office.context.mailbox.item.itemId;
            }
        } catch (e) {
            // Office.js not available, try AppState
            if (AppState.currentEmail && AppState.currentEmail.id) {
                currentEmailId = AppState.currentEmail.id;
            }
        }
        
        if (!currentEmailId || !AuthService.isSignedIn()) {
            return null;
        }
        
        // Get email with folder information
        const emailRequest = AuthService.graphRequest(
            `/me/messages/${currentEmailId}?$select=id,parentFolderId`
        );
        const email = await Helpers.withTimeout(
            emailRequest,
            5000,
            'Timeout getting email folder info'
        );
        
        if (!email.parentFolderId) {
            return null;
        }
        
        // Get the current folder (should be Quotes if we're in the right place)
        const currentFolderRequest = AuthService.graphRequest(
            `/me/mailFolders/${email.parentFolderId}?$select=id,displayName,parentFolderId`
        );
        const currentFolder = await Helpers.withTimeout(
            currentFolderRequest,
            5000,
            'Timeout getting current folder info'
        );
        
        if (!currentFolder) {
            return null;
        }
        
        // Check if we're in a Quotes folder
        if (currentFolder.displayName && currentFolder.displayName.toLowerCase() === 'quotes') {
            // Get the parent folder (should be the material folder MAT-XXXXX)
            if (currentFolder.parentFolderId) {
                const parentFolderRequest = AuthService.graphRequest(
                    `/me/mailFolders/${currentFolder.parentFolderId}?$select=id,displayName`
                );
                const parentFolder = await Helpers.withTimeout(
                    parentFolderRequest,
                    5000,
                    'Timeout getting parent folder info'
                );
                
                if (parentFolder && parentFolder.displayName && /^MAT-\d+$/i.test(parentFolder.displayName)) {
                    return parentFolder.displayName.toUpperCase();
                }
            }
        }
        
        // If not in Quotes folder, walk up the tree to find material folder
        let currentFolderId = email.parentFolderId;
        const maxDepth = 5;
        let depth = 0;
        
        while (currentFolderId && depth < maxDepth) {
            const folderRequest = AuthService.graphRequest(
                `/me/mailFolders/${currentFolderId}?$select=id,displayName,parentFolderId`
            );
            const folder = await Helpers.withTimeout(
                folderRequest,
                3000,
                'Timeout walking folder tree'
            );
            
            if (!folder) break;
            
            // Check if this is a material folder (MAT-XXXXX)
            if (folder.displayName && /^MAT-\d+$/i.test(folder.displayName)) {
                return folder.displayName.toUpperCase();
            }
            
            // Move up to parent folder
            if (!folder.parentFolderId || folder.parentFolderId === 'msgfolderroot') {
                break;
            }
            currentFolderId = folder.parentFolderId;
            depth++;
        }
        
        return null;
    } catch (error) {
        console.error('Error getting material code from email context:', error);
        return null;
    }
}

/**
 * Load all quotes from the Quotes folder for the current material
 * Automatically detects material code from email's folder context
 */
async function loadAllQuotesFromFolder() {
    const container = document.getElementById('quotes-container');
    if (!container) return;
    
    try {
        Helpers.showLoading('Loading quotes from folders...');
        container.innerHTML = '<div class="loading-indicator"><div class="spinner-small"></div><span>Detecting material from email context...</span></div>';
        
        if (!AuthService.isSignedIn()) {
            container.innerHTML = '<p class="placeholder-text">Please sign in to view quotes</p>';
            Helpers.hideLoading();
            return;
        }
        
        // Try to get material code from email's folder context first
        let materialCode = await getMaterialCodeFromEmailContext();
        
        // Fallback: Get material code from selected PR if available
        if (!materialCode && AppState.selectedPR) {
            materialCode = Helpers.extractMaterialCode(AppState.selectedPR);
        }
        
        // Fallback: Try to extract from email subject
        if (!materialCode) {
            try {
                if (Office.context.mailbox && Office.context.mailbox.item) {
                    const subject = Office.context.mailbox.item.subject || '';
                    const match = subject.match(/MAT-\d+/i);
                    if (match) {
                        materialCode = match[0].toUpperCase();
                    }
                }
            } catch (e) {
                // Office.js not available
            }
        }
        
        if (!materialCode) {
            container.innerHTML = '<p class="placeholder-text">Unable to detect material code. Please open an email from a Quotes folder or select a Purchase Requisition.</p>';
            Helpers.hideLoading();
            return;
        }
        
        // Find the Quotes folder for this specific material (MAT-XXXXX/Quotes)
        container.innerHTML = `<div class="loading-indicator"><div class="spinner-small"></div><span>Finding Quotes folder for ${materialCode}...</span></div>`;
        const quotesFolder = await findMaterialQuotesFolder(materialCode);
        
        if (!quotesFolder) {
            container.innerHTML = `<p class="placeholder-text">No Quotes folder found for ${materialCode}. Quotes will appear here once suppliers respond.</p>`;
            Helpers.hideLoading();
            return;
        }
        
        // Get all emails from the Quotes folder for this material
        container.innerHTML = `<div class="loading-indicator"><div class="spinner-small"></div><span>Loading emails from Quotes folder...</span></div>`;
        let allEmails = [];
        try {
            allEmails = await getEmailsByFolderId(quotesFolder.id, {
                top: 100,
                select: ['id', 'subject', 'from', 'body', 'receivedDateTime'],
                orderBy: 'receivedDateTime desc'
            });
        } catch (error) {
            console.error(`Error getting emails from Quotes folder for ${materialCode}:`, error);
            container.innerHTML = `<p class="placeholder-text">Error loading emails: ${Helpers.escapeHtml(error.message)}</p>`;
            Helpers.hideLoading();
            return;
        }
        
        if (allEmails.length === 0) {
            container.innerHTML = '<p class="placeholder-text">No quotes found in Quotes folders</p>';
            Helpers.hideLoading();
            return;
        }
        
        // Extract quote information from emails in batches (5 at a time)
        container.innerHTML = `<div class="loading-indicator"><div class="spinner-small"></div><span>Extracting quote data from ${allEmails.length} email(s)...</span></div>`;
        const quotes = [];
        const batchSize = 5;
        
        for (let i = 0; i < allEmails.length; i += batchSize) {
            const batch = allEmails.slice(i, i + batchSize);
            const batchNumber = Math.floor(i / batchSize) + 1;
            const totalBatches = Math.ceil(allEmails.length / batchSize);
            
            // Update progress
            if (allEmails.length > batchSize) {
                container.innerHTML = `<div class="loading-indicator"><div class="spinner-small"></div><span>Processing batch ${batchNumber} of ${totalBatches} (${i + 1}-${Math.min(i + batchSize, allEmails.length)} of ${allEmails.length})...</span></div>`;
            }
            
            // Process batch in parallel
            const batchPromises = batch.map(async (email) => {
                try {
                    const quote = await extractQuoteFromEmail(email);
                    return quote;
                } catch (error) {
                    console.error(`Error extracting quote from email ${email.id}:`, error);
                    // Return minimal quote info instead of failing
                    return {
                        supplier_name: email.from?.emailAddress?.name || email.from?.emailAddress?.address || 'Unknown',
                        supplier_email: email.from?.emailAddress?.address || '',
                        price: null,
                        unit_price: null,
                        total_price: null,
                        lead_time: null,
                        delivery_time: null,
                        validity: null,
                        payment_terms: null,
                        quote_date: email.receivedDateTime,
                        status: 'Received',
                        currency: 'USD',
                        email_id: email.id,
                        email_subject: email.subject
                    };
                }
            });
            
            const batchResults = await Promise.all(batchPromises);
            batchResults.forEach(quote => {
                if (quote) {
                    quotes.push(quote);
                }
            });
        }
        
        // Render comparison table
        renderQuoteComparison(quotes);
        Helpers.hideLoading();
        
    } catch (error) {
        console.error('Error loading quotes:', error);
        container.innerHTML = '<p class="placeholder-text">Error loading quotes: ' + Helpers.escapeHtml(error.message) + '</p>';
        Helpers.hideLoading();
    }
}

/**
 * Find the Quotes folder for a specific material (e.g., MAT-12345/Quotes)
 * Returns folder ID directly (no path resolution needed)
 * @param {string} materialCode - The material code (e.g., "MAT-12345")
 * @returns {Promise<Object|null>} Folder object with {id, name} or null if not found
 */
async function findMaterialQuotesFolder(materialCode) {
    try {
        // Get all mail folders with timeout
        const folderRequest = AuthService.graphRequest('/me/mailFolders?$top=500');
        const response = await Helpers.withTimeout(
            folderRequest,
            5000,
            'Timeout while fetching folders'
        );
        const allFolders = response.value || [];
        
        // Find the material folder (MAT-XXXXX)
        const materialFolder = allFolders.find(folder => 
            folder.displayName && folder.displayName.toUpperCase() === materialCode.toUpperCase()
        );
        
        if (!materialFolder) {
            console.log(`Material folder ${materialCode} not found`);
            return null;
        }
        
        // Get child folders of the material folder
        try {
            const childrenRequest = AuthService.graphRequest(
                `/me/mailFolders/${materialFolder.id}/childFolders?$top=100`
            );
            const children = await Helpers.withTimeout(
                childrenRequest,
                5000,
                `Timeout checking child folders of ${materialCode}`
            );
            
            if (children.value) {
                // Find the Quotes folder
                const quotesFolder = children.value.find(child =>
                    child.displayName && child.displayName.toLowerCase() === 'quotes'
                );
                
                if (quotesFolder) {
                    return {
                        id: quotesFolder.id,
                        name: quotesFolder.displayName
                    };
                }
            }
        } catch (error) {
            console.error(`Error getting child folders for ${materialCode}:`, error);
            return null;
        }
        
        return null;
    } catch (error) {
        console.error('Error finding Quotes folder:', error);
        return null;
    }
}

/**
 * Get emails directly from a folder by ID (no path conversion needed)
 * @param {string} folderId - The folder ID
 * @param {Object} options - Options for fetching emails
 * @returns {Promise<Array>} Array of email objects
 */
async function getEmailsByFolderId(folderId, options = {}) {
    try {
        let endpoint = `/me/mailFolders/${folderId}/messages`;
        const params = [];

        if (options.top) {
            params.push(`$top=${options.top}`);
        }
        if (options.select) {
            params.push(`$select=${options.select.join(',')}`);
        }
        if (options.orderBy) {
            params.push(`$orderby=${options.orderBy}`);
        }

        if (params.length > 0) {
            endpoint += '?' + params.join('&');
        }

        const emailRequest = AuthService.graphRequest(endpoint);
        const response = await Helpers.withTimeout(
            emailRequest,
            10000,
            `Timeout fetching emails from folder ${folderId}`
        );
        
        return response.value || [];
    } catch (error) {
        console.error(`Error getting emails from folder ${folderId}:`, error);
        return [];
    }
}

/**
 * Extract quote information from an email
 * Optimized to only fetch body if missing, with timeout protection
 */
async function extractQuoteFromEmail(email) {
    try {
        // Get full email body if not already available (with timeout)
        let emailBody = email.body?.content || '';
        if (!emailBody && email.id) {
            try {
                const emailRequest = EmailOperations.getEmailById(email.id);
                const fullEmail = await Helpers.withTimeout(
                    emailRequest,
                    5000,
                    `Timeout fetching email body for ${email.id}`
                );
                emailBody = fullEmail.body?.content || '';
            } catch (error) {
                console.warn(`Could not fetch full email body for ${email.id}:`, error.message);
                // Continue with empty body - extraction will work with what we have
                emailBody = '';
            }
        }
        
        const bodyText = Helpers.stripHtml(emailBody);
        const supplierName = email.from?.emailAddress?.name || email.from?.emailAddress?.address || 'Unknown';
        const supplierEmail = email.from?.emailAddress?.address || '';
        
        // Try to extract quote data using patterns
        const quote = {
            supplier_name: supplierName,
            supplier_email: supplierEmail,
            email_id: email.id,
            email_subject: email.subject,
            quote_date: email.receivedDateTime,
            status: 'Received',
            currency: 'USD'
        };
        
        // Extract prices
        // Pattern: $X.XX, USD X.XX, Price: X, Unit Price: X, Total: X
        const pricePatterns = [
            /\$[\d,]+\.?\d*/g,
            /USD\s*[\d,]+\.?\d*/gi,
            /(?:unit\s*)?price[:\s]*\$?[\d,]+\.?\d*/gi,
            /total[:\s]*\$?[\d,]+\.?\d*/gi,
            /[\d,]+\.?\d*\s*(?:USD|dollars?)/gi
        ];
        
        let prices = [];
        for (const pattern of pricePatterns) {
            const matches = bodyText.match(pattern);
            if (matches) {
                prices.push(...matches);
            }
        }
        
        // Extract numeric values from prices
        const numericPrices = prices.map(p => {
            const num = parseFloat(p.replace(/[^0-9.]/g, ''));
            return isNaN(num) ? null : num;
        }).filter(p => p !== null && p > 0);
        
        // Try to identify unit price vs total price from context
        const unitPriceMatch = bodyText.match(/(?:unit\s*price|price\s*per\s*unit)[:\s]*\$?([\d,]+\.?\d*)/i);
        const totalPriceMatch = bodyText.match(/(?:total\s*price|total\s*amount|grand\s*total)[:\s]*\$?([\d,]+\.?\d*)/i);
        
        if (unitPriceMatch) {
            quote.unit_price = parseFloat(unitPriceMatch[1].replace(/[^0-9.]/g, ''));
        }
        if (totalPriceMatch) {
            quote.total_price = parseFloat(totalPriceMatch[1].replace(/[^0-9.]/g, ''));
        }
        
        // If we found prices but didn't identify unit/total, make educated guesses
        if (numericPrices.length > 0) {
            if (!quote.unit_price && !quote.total_price) {
                // If multiple prices, assume largest is total, smallest is unit
                if (numericPrices.length > 1) {
                    quote.total_price = Math.max(...numericPrices);
                    quote.unit_price = Math.min(...numericPrices);
                } else {
                    // Single price - could be either, assume it's total
                    quote.total_price = numericPrices[0];
                    quote.unit_price = numericPrices[0];
                }
            } else if (!quote.total_price && quote.unit_price) {
                // Have unit price, try to find total
                quote.total_price = numericPrices.find(p => p > quote.unit_price) || quote.unit_price;
            } else if (!quote.unit_price && quote.total_price) {
                // Have total price, try to find unit
                quote.unit_price = numericPrices.find(p => p < quote.total_price) || quote.total_price;
            }
            quote.price = quote.total_price || quote.unit_price || numericPrices[0];
        }
        
        // Extract lead time / delivery time
        const leadTimePatterns = [
            /(?:lead\s*time|delivery\s*time)[:\s]*(\d+\s*(?:weeks?|days?|months?))/gi,
            /(\d+\s*(?:weeks?|days?|months?))\s*(?:lead|delivery)/gi,
            /delivery[:\s]*(\d+\s*(?:weeks?|days?|months?))/gi
        ];
        
        for (const pattern of leadTimePatterns) {
            const match = bodyText.match(pattern);
            if (match && match[1]) {
                quote.lead_time = match[1].trim();
                quote.delivery_time = match[1].trim();
                break;
            }
        }
        
        // Extract validity
        const validityPatterns = [
            /valid(?:ity)?[:\s]*(?:for\s*)?(\d+\s*(?:days?|weeks?|months?))/gi,
            /valid\s*(?:for|until)[:\s]*(\d+\s*(?:days?|weeks?|months?))/gi,
            /expires?\s*(?:in|on)[:\s]*(\d+\s*(?:days?|weeks?|months?))/gi
        ];
        
        for (const pattern of validityPatterns) {
            const match = bodyText.match(pattern);
            if (match && match[1]) {
                quote.validity = match[1].trim();
                break;
            }
        }
        
        // Extract payment terms
        const paymentPatterns = [
            /payment\s*terms?[:\s]*(net\s*\d+|[\w\s]+)/gi,
            /net\s*\d+/gi,
            /terms?[:\s]*(net\s*\d+|[\w\s]+)/gi
        ];
        
        for (const pattern of paymentPatterns) {
            const match = bodyText.match(pattern);
            if (match && match[1] || match[0]) {
                quote.payment_terms = (match[1] || match[0]).trim();
                break;
            }
        }
        
        return quote;
    } catch (error) {
        console.error('Error extracting quote from email:', error);
        return null;
    }
}

async function handleRFQSelect(event) {
    const rfqId = event.target.value;
    
    if (!rfqId) {
        document.getElementById('quotes-container').innerHTML = 
            '<p class="placeholder-text">Select an RFQ to view and compare quotes</p>';
        return;
    }
    
    try {
        Helpers.showLoading('Loading quotes...');
        
        const quotes = await ApiClient.getQuotes(rfqId);
        renderQuoteComparison(quotes);
    } catch (error) {
        Helpers.showError('Failed to load quotes: ' + error.message);
    } finally {
        Helpers.hideLoading();
    }
}

function renderQuoteComparison(quotes) {
    const container = document.getElementById('quotes-container');
    if (!container) return;
    
    Helpers.clearChildren(container);
    
    if (quotes.length === 0) {
        container.innerHTML = '<p class="placeholder-text">No quotes found in Quotes folders</p>';
        Helpers.hideElement(document.getElementById('quote-summary'));
        return;
    }
    
    // Calculate best prices and statistics (only from quotes with valid prices)
    const prices = quotes
        .map(q => {
            // Try unit_price first, then total_price, then price
            const unitPrice = q.unit_price ? parseFloat(q.unit_price) : null;
            const totalPrice = q.total_price ? parseFloat(q.total_price) : null;
            const price = q.price ? parseFloat(q.price) : null;
            
            // Prefer unit price for comparison, fallback to total or price
            const comparisonPrice = unitPrice || totalPrice || price;
            return comparisonPrice && comparisonPrice > 0 ? comparisonPrice : null;
        })
        .filter(p => p !== null && p > 0);
    
    const lowestPrice = prices.length > 0 ? Math.min(...prices) : null;
    const highestPrice = prices.length > 0 ? Math.max(...prices) : null;
    const averagePrice = prices.length > 0 ? prices.reduce((a, b) => a + b, 0) / prices.length : null;
    
    // Find fastest delivery (assuming delivery_time is in a comparable format)
    const quotesWithDelivery = quotes.filter(q => q.delivery_time);
    const fastestDelivery = quotesWithDelivery.length > 0 
        ? quotesWithDelivery.reduce((fastest, current) => {
            // Simple comparison - in production, parse delivery times properly
            return current;
        }, quotesWithDelivery[0])
        : null;
    
    // Create comparison table
    const tableWrapper = document.createElement('div');
    tableWrapper.className = 'quote-comparison-table-wrapper';
    
    const table = document.createElement('table');
    table.className = 'comparison-table';
    
    // Table header
    const thead = document.createElement('thead');
    thead.className = 'comparison-header';
    thead.innerHTML = `
        <tr>
            <th>Supplier</th>
            <th>Unit Price</th>
            <th>Total Price</th>
            <th>Lead Time</th>
            <th>Validity</th>
            <th>Payment Terms</th>
            <th>Quote Date</th>
            <th>Status</th>
        </tr>
    `;
    
    // Calculate lowest prices for highlighting (before creating rows)
    const unitPrices = quotes
        .map(q => {
            const up = parseFloat(q.unit_price) || parseFloat(q.total_price) || parseFloat(q.price) || 0;
            return up > 0 ? up : null;
        })
        .filter(p => p !== null && p > 0);
    const totalPrices = quotes
        .map(q => {
            const tp = parseFloat(q.total_price) || parseFloat(q.price) || 0;
            return tp > 0 ? tp : null;
        })
        .filter(p => p !== null && p > 0);
    
    const lowestUnitPrice = unitPrices.length > 0 ? Math.min(...unitPrices) : null;
    const lowestTotalPrice = totalPrices.length > 0 ? Math.min(...totalPrices) : null;
    
    // Table body
    const tbody = document.createElement('tbody');
    
    quotes.forEach((quote, index) => {
        const row = document.createElement('tr');
        row.className = 'comparison-row';
        if (index % 2 === 0) {
            row.classList.add('even-row');
        }
        
        // Handle quotes with minimal information gracefully
        const unitPrice = quote.unit_price ? parseFloat(quote.unit_price) : null;
        const totalPrice = quote.total_price ? parseFloat(quote.total_price) : null;
        const price = quote.price ? parseFloat(quote.price) : null;
        
        // For comparison, use numeric values
        const unitPriceNum = unitPrice !== null && !isNaN(unitPrice) ? unitPrice : (totalPrice !== null && !isNaN(totalPrice) ? totalPrice : (price !== null && !isNaN(price) ? price : 0));
        const totalPriceNum = totalPrice !== null && !isNaN(totalPrice) ? totalPrice : (price !== null && !isNaN(price) ? price : (unitPrice !== null && !isNaN(unitPrice) ? unitPrice : 0));
        
        // Check if this is the lowest
        const isLowestUnit = unitPriceNum > 0 && lowestUnitPrice !== null && unitPriceNum === lowestUnitPrice;
        const isLowestTotal = totalPriceNum > 0 && lowestTotalPrice !== null && totalPriceNum === lowestTotalPrice;
        
        row.innerHTML = `
            <td class="comparison-cell supplier-cell">
                <strong>${Helpers.escapeHtml(quote.supplier_name || 'Unknown')}</strong>
            </td>
            <td class="comparison-cell price-cell ${isLowestUnit ? 'best-price' : ''}">
                ${unitPriceNum > 0 ? `
                    <span class="price-value">${Helpers.formatCurrency(unitPriceNum, quote.currency || 'USD')}</span>
                    ${isLowestUnit ? '<span class="price-badge">Best</span>' : ''}
                ` : '<span class="no-data">-</span>'}
            </td>
            <td class="comparison-cell price-cell ${isLowestTotal ? 'best-price' : ''}">
                ${totalPriceNum > 0 ? `
                    <span class="price-value">${Helpers.formatCurrency(totalPriceNum, quote.currency || 'USD')}</span>
                    ${isLowestTotal ? '<span class="price-badge">Best</span>' : ''}
                ` : '<span class="no-data">-</span>'}
            </td>
            <td class="comparison-cell">
                ${quote.lead_time || quote.delivery_time || '<span class="no-data">-</span>'}
            </td>
            <td class="comparison-cell">
                ${quote.validity || quote.validity_period || '<span class="no-data">-</span>'}
            </td>
            <td class="comparison-cell">
                ${quote.payment_terms || '<span class="no-data">-</span>'}
            </td>
            <td class="comparison-cell">
                ${quote.quote_date ? Helpers.formatDate(quote.quote_date) : '<span class="no-data">-</span>'}
            </td>
            <td class="comparison-cell status-cell">
                <span class="quote-status-badge ${(quote.status || '').toLowerCase().replace(/\s+/g, '-')}">
                    ${Helpers.escapeHtml(quote.status || 'Pending')}
                </span>
            </td>
        `;
        
        tbody.appendChild(row);
    });
    
    table.appendChild(thead);
    table.appendChild(tbody);
    tableWrapper.appendChild(table);
    container.appendChild(tableWrapper);
    
    // Enhanced summary
    const summaryContainer = document.getElementById('summary-content');
    const summarySection = document.getElementById('quote-summary');
    
    if (quotes.length > 0) {
        if (summaryContainer) {
            const lowestQuote = quotes.find(q => {
                const price = parseFloat(q.unit_price) || parseFloat(q.total_price) || parseFloat(q.price) || 0;
                return price > 0 && price === lowestPrice;
            }) || quotes[0];
            
            summaryContainer.innerHTML = `
                <div class="summary-card">
                    <div class="summary-label">Lowest Price</div>
                    <div class="summary-value highlight">
                        ${lowestPrice ? Helpers.formatCurrency(lowestPrice, lowestQuote.currency || 'USD') : 'N/A'}
                    </div>
                    <div class="summary-subtext">${Helpers.escapeHtml(lowestQuote.supplier_name || '')}</div>
                </div>
                ${averagePrice ? `
                <div class="summary-card">
                    <div class="summary-label">Average Price</div>
                    <div class="summary-value">
                        ${Helpers.formatCurrency(averagePrice, lowestQuote.currency || 'USD')}
                    </div>
                </div>
                ` : ''}
                ${lowestPrice && highestPrice ? `
                <div class="summary-card">
                    <div class="summary-label">Price Range</div>
                    <div class="summary-value">
                        ${Helpers.formatCurrency(lowestPrice, lowestQuote.currency || 'USD')} - 
                        ${Helpers.formatCurrency(highestPrice, lowestQuote.currency || 'USD')}
                    </div>
                </div>
                ` : ''}
                <div class="summary-card">
                    <div class="summary-label">Total Quotes</div>
                    <div class="summary-value">${quotes.length}</div>
                </div>
                ${fastestDelivery ? `
                <div class="summary-card">
                    <div class="summary-label">Fastest Delivery</div>
                    <div class="summary-value">${Helpers.escapeHtml(fastestDelivery.delivery_time || fastestDelivery.lead_time || 'N/A')}</div>
                    <div class="summary-subtext">${Helpers.escapeHtml(fastestDelivery.supplier_name || '')}</div>
                </div>
                ` : ''}
            `;
        }
        if (summarySection) {
            Helpers.showElement(summarySection);
        }
    } else {
        if (summarySection) {
            Helpers.hideElement(summarySection);
        }
    }
}

// ==================== QUOTE COMPARISON MODAL ====================

// State for modal
let modalQuotesState = {
    allQuotes: [],
    filteredQuotes: [],
    sortBy: 'unit_price_asc',
    filters: {
        search: '',
        bestPrice: false,
        fastestDelivery: false
    }
};

/**
 * Open the quote comparison modal
 */
async function openQuoteComparisonModal() {
    const modal = document.getElementById('quote-comparison-modal');
    if (!modal) {
        console.error('Quote comparison modal not found');
        return;
    }
    
    // Show modal
    modal.classList.remove('hidden');
    
    // Show loading state
    const loadingEl = document.getElementById('quote-comparison-loading');
    const tableWrapper = document.getElementById('quote-comparison-table-wrapper');
    const emptyState = document.getElementById('quote-comparison-empty');
    
    if (loadingEl) Helpers.showElement(loadingEl);
    if (tableWrapper) Helpers.hideElement(tableWrapper);
    if (emptyState) Helpers.hideElement(emptyState);
    
    try {
        // Load all quotes for the modal
        const quotes = await getAllQuotesForModal();
        
        // Store in state
        modalQuotesState.allQuotes = quotes;
        modalQuotesState.filteredQuotes = [...quotes];
        
        // Apply initial sort
        const [sortField, sortDirection] = modalQuotesState.sortBy.split('_');
        const sortedQuotes = sortQuotes(quotes, sortField, sortDirection === 'desc' ? 'desc' : 'asc');
        modalQuotesState.filteredQuotes = sortedQuotes;
        
        // Render modal
        renderQuoteComparisonModal(sortedQuotes);
        
    } catch (error) {
        console.error('Error opening quote comparison modal:', error);
        if (loadingEl) {
            loadingEl.innerHTML = `<div class="error-message">Error loading quotes: ${Helpers.escapeHtml(error.message)}</div>`;
        }
        Helpers.showError('Failed to load quotes: ' + error.message);
    }
}

/**
 * Get all quotes for the modal (reuse existing logic)
 */
async function getAllQuotesForModal() {
    const quotes = [];
    
    if (!AuthService.isSignedIn()) {
        throw new Error('Please sign in to view quotes');
    }
    
    try {
        // Get all mail folders
        const foldersResponse = await AuthService.graphRequest('/me/mailFolders?$top=500');
        const allFolders = foldersResponse.value || [];
        
        // Find all material folders (MAT-XXXXX)
        const materialFolders = allFolders.filter(folder => {
            const name = folder.displayName || '';
            return /^MAT-\d+/i.test(name);
        });
        
        // For each material folder, find its Quotes subfolder
        for (const materialFolder of materialFolders) {
            try {
                const childrenResponse = await AuthService.graphRequest(
                    `/me/mailFolders/${materialFolder.id}/childFolders?$top=100`
                );
                
                if (childrenResponse.value) {
                    const quotesFolder = childrenResponse.value.find(child =>
                        child.displayName && child.displayName.toLowerCase() === 'quotes'
                    );
                    
                    if (quotesFolder) {
                        // Get emails from this Quotes folder
                        const emailsResponse = await AuthService.graphRequest(
                            `/me/mailFolders/${quotesFolder.id}/messages?$top=100&$select=id,subject,from,body,receivedDateTime`
                        );
                        
                        if (emailsResponse.value) {
                            for (const email of emailsResponse.value) {
                                // Extract quote data from email
                                const quote = extractQuoteFromEmail(email, materialFolder.displayName);
                                if (quote) {
                                    quotes.push(quote);
                                }
                            }
                        }
                    }
                }
            } catch (folderError) {
                console.warn(`Error processing folder ${materialFolder.displayName}:`, folderError);
                // Continue with other folders
            }
        }
    } catch (error) {
        console.error('Error loading quotes for modal:', error);
        throw error;
    }
    
    return quotes;
}

/**
 * Extract quote data from email
 */
function extractQuoteFromEmail(email, materialCode) {
    try {
        const bodyText = email.body?.content || '';
        const fromEmail = email.from?.emailAddress?.address || '';
        const fromName = email.from?.emailAddress?.name || '';
        
        // Try to extract price information
        const priceMatch = bodyText.match(/(?:unit\s*price|price)[:\s]*\$?([\d,]+\.?\d*)/i);
        const totalMatch = bodyText.match(/(?:total\s*price|total)[:\s]*\$?([\d,]+\.?\d*)/i);
        const deliveryMatch = bodyText.match(/(?:delivery|lead\s*time)[:\s]*([^\n]+)/i);
        const validityMatch = bodyText.match(/(?:validity|valid)[:\s]*([^\n]+)/i);
        const termsMatch = bodyText.match(/(?:payment\s*terms|terms)[:\s]*([^\n]+)/i);
        
        const unitPrice = priceMatch ? parseFloat(priceMatch[1].replace(/,/g, '')) : null;
        const totalPrice = totalMatch ? parseFloat(totalMatch[1].replace(/,/g, '')) : null;
        const deliveryTime = deliveryMatch ? deliveryMatch[1].trim() : null;
        const validity = validityMatch ? validityMatch[1].trim() : null;
        const paymentTerms = termsMatch ? termsMatch[1].trim() : null;
        
        return {
            supplier_name: fromName || fromEmail.split('@')[0],
            supplier_email: fromEmail,
            unit_price: unitPrice,
            total_price: totalPrice,
            price: unitPrice || totalPrice,
            lead_time: deliveryTime,
            delivery_time: deliveryTime,
            validity: validity,
            validity_period: validity,
            payment_terms: paymentTerms,
            quote_date: email.receivedDateTime,
            currency: 'USD',
            status: 'Received',
            material_code: materialCode,
            email_id: email.id,
            subject: email.subject
        };
    } catch (error) {
        console.error('Error extracting quote from email:', error);
        return null;
    }
}

/**
 * Render quote comparison in the modal
 */
function renderQuoteComparisonModal(quotes) {
    const loadingEl = document.getElementById('quote-comparison-loading');
    const tableWrapper = document.getElementById('quote-comparison-table-wrapper');
    const emptyState = document.getElementById('quote-comparison-empty');
    const summaryCards = document.getElementById('quote-summary-cards');
    const countDisplay = document.getElementById('quote-count-display');
    
    // Hide loading
    if (loadingEl) Helpers.hideElement(loadingEl);
    
    // Update quote count
    if (countDisplay) {
        countDisplay.textContent = `${quotes.length} quote${quotes.length !== 1 ? 's' : ''}`;
    }
    
    if (quotes.length === 0) {
        if (tableWrapper) Helpers.hideElement(tableWrapper);
        if (emptyState) Helpers.showElement(emptyState);
        if (summaryCards) summaryCards.innerHTML = '';
        return;
    }
    
    if (emptyState) Helpers.hideElement(emptyState);
    if (tableWrapper) Helpers.showElement(tableWrapper);
    
    // Render summary cards
    renderSummaryCards(quotes, summaryCards);
    
    // Render comparison table
    renderModalComparisonTable(quotes, tableWrapper);
}

/**
 * Render summary cards
 */
function renderSummaryCards(quotes, container) {
    if (!container) return;
    
    // Calculate statistics
    const prices = quotes
        .map(q => {
            const price = parseFloat(q.unit_price) || parseFloat(q.total_price) || parseFloat(q.price) || 0;
            return price > 0 ? price : null;
        })
        .filter(p => p !== null && p > 0);
    
    const lowestPrice = prices.length > 0 ? Math.min(...prices) : null;
    const highestPrice = prices.length > 0 ? Math.max(...prices) : null;
    const averagePrice = prices.length > 0 ? prices.reduce((a, b) => a + b, 0) / prices.length : null;
    
    const lowestQuote = quotes.find(q => {
        const price = parseFloat(q.unit_price) || parseFloat(q.total_price) || parseFloat(q.price) || 0;
        return price > 0 && price === lowestPrice;
    }) || quotes[0];
    
    // Find fastest delivery
    const quotesWithDelivery = quotes.filter(q => q.delivery_time || q.lead_time);
    const fastestQuote = quotesWithDelivery.length > 0 ? quotesWithDelivery[0] : null;
    
    container.innerHTML = `
        <div class="summary-card card-total">
            <div class="summary-card-content">
                <div class="summary-card-label">Total Quotes</div>
                <div class="summary-card-value">${quotes.length}</div>
            </div>
        </div>
        <div class="summary-card card-best">
            <div class="summary-card-content">
                <div class="summary-card-label">Best Price</div>
                <div class="summary-card-value">${lowestPrice ? Helpers.formatCurrency(lowestPrice, lowestQuote.currency || 'USD') : 'N/A'}</div>
                <div class="summary-card-subtext">${Helpers.escapeHtml(lowestQuote.supplier_name || '')}</div>
            </div>
        </div>
        <div class="summary-card card-average">
            <div class="summary-card-content">
                <div class="summary-card-label">Average Price</div>
                <div class="summary-card-value">${averagePrice ? Helpers.formatCurrency(averagePrice, lowestQuote.currency || 'USD') : 'N/A'}</div>
            </div>
        </div>
        <div class="summary-card card-fastest">
            <div class="summary-card-content">
                <div class="summary-card-label">Fastest Delivery</div>
                <div class="summary-card-value">${fastestQuote ? Helpers.escapeHtml(fastestQuote.delivery_time || fastestQuote.lead_time || 'N/A') : 'N/A'}</div>
                <div class="summary-card-subtext">${fastestQuote ? Helpers.escapeHtml(fastestQuote.supplier_name || '') : ''}</div>
            </div>
        </div>
    `;
}

/**
 * Render comparison table in modal
 */
function renderModalComparisonTable(quotes, container) {
    if (!container) return;
    
    // Calculate best prices for highlighting
    const unitPrices = quotes
        .map(q => {
            const up = parseFloat(q.unit_price) || parseFloat(q.total_price) || parseFloat(q.price) || 0;
            return up > 0 ? up : null;
        })
        .filter(p => p !== null && p > 0);
    
    const totalPrices = quotes
        .map(q => {
            const tp = parseFloat(q.total_price) || parseFloat(q.price) || 0;
            return tp > 0 ? tp : null;
        })
        .filter(p => p !== null && p > 0);
    
    const lowestUnitPrice = unitPrices.length > 0 ? Math.min(...unitPrices) : null;
    const lowestTotalPrice = totalPrices.length > 0 ? Math.min(...totalPrices) : null;
    
    // Find fastest delivery time
    const deliveryTimes = quotes
        .map(q => q.delivery_time || q.lead_time)
        .filter(t => t && t.trim().length > 0);
    
    const fastestDelivery = deliveryTimes.length > 0 ? deliveryTimes[0] : null;
    
    // Create table
    const table = document.createElement('table');
    table.className = 'quote-comparison-table';
    
    // Table header
    const thead = document.createElement('thead');
    thead.innerHTML = `
        <tr>
            <th>Supplier</th>
            <th>Unit Price</th>
            <th>Total Price</th>
            <th>Lead Time</th>
            <th>Validity</th>
            <th>Payment Terms</th>
            <th>Quote Date</th>
            <th>Actions</th>
        </tr>
    `;
    
    // Table body
    const tbody = document.createElement('tbody');
    
    quotes.forEach((quote, index) => {
        const row = document.createElement('tr');
        row.className = 'comparison-row';
        row.dataset.quoteIndex = index;
        
        if (index % 2 === 0) {
            row.classList.add('even-row');
        }
        
        const unitPrice = parseFloat(quote.unit_price) || 0;
        const totalPrice = parseFloat(quote.total_price) || parseFloat(quote.price) || 0;
        const deliveryTime = quote.delivery_time || quote.lead_time || '';
        
        const isLowestUnit = unitPrice > 0 && lowestUnitPrice !== null && unitPrice === lowestUnitPrice;
        const isLowestTotal = totalPrice > 0 && lowestTotalPrice !== null && totalPrice === lowestTotalPrice;
        const isFastest = deliveryTime && deliveryTime === fastestDelivery;
        
        row.innerHTML = `
            <td class="supplier-cell">
                <strong>${Helpers.escapeHtml(quote.supplier_name || 'Unknown')}</strong>
                <div class="supplier-email">${Helpers.escapeHtml(quote.supplier_email || '')}</div>
            </td>
            <td class="price-cell ${isLowestUnit ? 'best-price' : ''}">
                ${unitPrice > 0 ? `
                    <span class="price-value">${Helpers.formatCurrency(unitPrice, quote.currency || 'USD')}</span>
                    ${isLowestUnit ? '<span class="best-badge">Best</span>' : ''}
                ` : '<span class="no-data">-</span>'}
            </td>
            <td class="price-cell ${isLowestTotal ? 'best-price' : ''}">
                ${totalPrice > 0 ? `
                    <span class="price-value">${Helpers.formatCurrency(totalPrice, quote.currency || 'USD')}</span>
                    ${isLowestTotal ? '<span class="best-badge">Best</span>' : ''}
                ` : '<span class="no-data">-</span>'}
            </td>
            <td class="delivery-cell ${isFastest ? 'fastest-delivery' : ''}">
                ${deliveryTime ? `
                    ${Helpers.escapeHtml(deliveryTime)}
                    ${isFastest ? '<span class="fastest-badge">Fastest</span>' : ''}
                ` : '<span class="no-data">-</span>'}
            </td>
            <td>${quote.validity || quote.validity_period || '<span class="no-data">-</span>'}</td>
            <td>${quote.payment_terms || '<span class="no-data">-</span>'}</td>
            <td>${quote.quote_date ? Helpers.formatDate(quote.quote_date) : '<span class="no-data">-</span>'}</td>
            <td class="actions-cell">
                <button class="ms-Button ms-Button--small accept-quote-row-btn" data-quote-index="${index}">
                    <span class="ms-Button-label">Accept</span>
                </button>
            </td>
        `;
        
        tbody.appendChild(row);
    });
    
    table.appendChild(thead);
    table.appendChild(tbody);
    
    // Clear and append
    Helpers.clearChildren(container);
    container.appendChild(table);
    
    // Attach event handlers for accept buttons
    container.querySelectorAll('.accept-quote-row-btn').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const index = parseInt(btn.dataset.quoteIndex);
            const quote = quotes[index];
            if (quote) {
                handleAcceptQuoteFromModal(quote);
            }
        });
    });
}

/**
 * Sort quotes based on criteria
 */
function sortQuotes(quotes, sortBy, direction = 'asc') {
    const sorted = [...quotes];
    
    sorted.sort((a, b) => {
        let aVal, bVal;
        
        switch (sortBy) {
            case 'unit_price':
                aVal = parseFloat(a.unit_price) || parseFloat(a.total_price) || parseFloat(a.price) || 0;
                bVal = parseFloat(b.unit_price) || parseFloat(b.total_price) || parseFloat(b.price) || 0;
                break;
            case 'total_price':
                aVal = parseFloat(a.total_price) || parseFloat(a.price) || 0;
                bVal = parseFloat(b.total_price) || parseFloat(b.price) || 0;
                break;
            case 'delivery':
                aVal = (a.delivery_time || a.lead_time || '').toLowerCase();
                bVal = (b.delivery_time || b.lead_time || '').toLowerCase();
                // For delivery, we want fastest first (shorter strings typically mean faster)
                return direction === 'asc' ? aVal.localeCompare(bVal) : bVal.localeCompare(aVal);
            case 'supplier':
                aVal = (a.supplier_name || '').toLowerCase();
                bVal = (b.supplier_name || '').toLowerCase();
                break;
            case 'date':
                aVal = new Date(a.quote_date || 0).getTime();
                bVal = new Date(b.quote_date || 0).getTime();
                break;
            default:
                return 0;
        }
        
        if (sortBy === 'delivery') {
            return 0; // Already handled above
        }
        
        if (typeof aVal === 'number' && typeof bVal === 'number') {
            return direction === 'asc' ? aVal - bVal : bVal - aVal;
        } else {
            return direction === 'asc' ? aVal.localeCompare(bVal) : bVal.localeCompare(aVal);
        }
    });
    
    return sorted;
}

/**
 * Filter quotes based on criteria
 */
function filterQuotes(quotes, filters) {
    let filtered = [...quotes];
    
    // Search filter
    if (filters.search && filters.search.trim()) {
        const searchLower = filters.search.toLowerCase();
        filtered = filtered.filter(q => {
            const supplierName = (q.supplier_name || '').toLowerCase();
            const supplierEmail = (q.supplier_email || '').toLowerCase();
            return supplierName.includes(searchLower) || supplierEmail.includes(searchLower);
        });
    }
    
    // Best price filter
    if (filters.bestPrice) {
        const prices = filtered
            .map(q => parseFloat(q.unit_price) || parseFloat(q.total_price) || parseFloat(q.price) || 0)
            .filter(p => p > 0);
        
        if (prices.length > 0) {
            const lowestPrice = Math.min(...prices);
            filtered = filtered.filter(q => {
                const price = parseFloat(q.unit_price) || parseFloat(q.total_price) || parseFloat(q.price) || 0;
                return price > 0 && price === lowestPrice;
            });
        }
    }
    
    // Fastest delivery filter
    if (filters.fastestDelivery) {
        const deliveryTimes = filtered
            .map(q => q.delivery_time || q.lead_time)
            .filter(t => t && t.trim().length > 0);
        
        if (deliveryTimes.length > 0) {
            // Simple approach: take the first one with delivery time
            // In production, you'd parse and compare delivery times properly
            const fastest = deliveryTimes[0];
            filtered = filtered.filter(q => {
                const delivery = q.delivery_time || q.lead_time;
                return delivery && delivery === fastest;
            });
        }
    }
    
    return filtered;
}

/**
 * Apply sorting and filtering to modal quotes
 */
function applyModalFiltersAndSort() {
    const { allQuotes, sortBy, filters } = modalQuotesState;
    
    // Parse sort criteria
    const [sortField, sortDirection] = sortBy.split('_');
    const direction = sortDirection === 'desc' ? 'desc' : 'asc';
    
    // Filter first
    let filtered = filterQuotes(allQuotes, filters);
    
    // Then sort
    filtered = sortQuotes(filtered, sortField, direction);
    
    // Update state
    modalQuotesState.filteredQuotes = filtered;
    
    // Re-render
    const tableWrapper = document.getElementById('quote-comparison-table-wrapper');
    renderModalComparisonTable(filtered, tableWrapper);
    
    // Update summary cards
    const summaryCards = document.getElementById('quote-summary-cards');
    renderSummaryCards(filtered, summaryCards);
    
    // Update count
    const countDisplay = document.getElementById('quote-count-display');
    if (countDisplay) {
        countDisplay.textContent = `${filtered.length} quote${filtered.length !== 1 ? 's' : ''}`;
    }
}

/**
 * Export quotes to CSV
 */
function exportQuotesToCSV(quotes) {
    if (quotes.length === 0) {
        Helpers.showError('No quotes to export');
        return;
    }
    
    // CSV header
    const headers = ['Supplier', 'Supplier Email', 'Unit Price', 'Total Price', 'Lead Time', 'Validity', 'Payment Terms', 'Quote Date', 'Status'];
    
    // CSV rows
    const rows = quotes.map(quote => {
        const unitPrice = parseFloat(quote.unit_price) || '';
        const totalPrice = parseFloat(quote.total_price) || parseFloat(quote.price) || '';
        return [
            quote.supplier_name || '',
            quote.supplier_email || '',
            unitPrice,
            totalPrice,
            quote.lead_time || quote.delivery_time || '',
            quote.validity || quote.validity_period || '',
            quote.payment_terms || '',
            quote.quote_date ? new Date(quote.quote_date).toLocaleDateString() : '',
            quote.status || ''
        ];
    });
    
    // Combine header and rows
    const csvContent = [
        headers.join(','),
        ...rows.map(row => row.map(cell => `"${String(cell).replace(/"/g, '""')}"`).join(','))
    ].join('\n');
    
    // Create download
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    link.setAttribute('href', url);
    link.setAttribute('download', `quote-comparison-${new Date().toISOString().split('T')[0]}.csv`);
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    
    Helpers.showSuccess('Quote comparison exported to CSV');
}

/**
 * Export quotes to PDF (using browser print)
 */
function exportQuotesToPDF(quotes) {
    if (quotes.length === 0) {
        Helpers.showError('No quotes to export');
        return;
    }
    
    // Create a printable version
    const printWindow = window.open('', '_blank');
    const tableWrapper = document.getElementById('quote-comparison-table-wrapper');
    const summaryCards = document.getElementById('quote-summary-cards');
    
    if (!tableWrapper || !summaryCards) {
        Helpers.showError('Could not generate PDF');
        return;
    }
    
    // Get table HTML
    const tableHTML = tableWrapper.innerHTML;
    const summaryHTML = summaryCards.innerHTML;
    
    // Create print document
    printWindow.document.write(`
        <!DOCTYPE html>
        <html>
        <head>
            <title>Quote Comparison</title>
            <style>
                body { font-family: Arial, sans-serif; padding: 20px; }
                h1 { color: #0d3d61; }
                .summary-cards-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 15px; margin-bottom: 30px; }
                .summary-card { border: 1px solid #ddd; padding: 15px; border-radius: 4px; }
                .summary-card-label { font-size: 12px; color: #666; }
                .summary-card-value { font-size: 24px; font-weight: bold; margin: 5px 0; }
                table { width: 100%; border-collapse: collapse; margin-top: 20px; }
                th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
                th { background-color: #0d3d61; color: white; }
                .best-price { background-color: #d4edda !important; }
                .fastest-delivery { background-color: #d1ecf1 !important; }
                @media print {
                    .no-print { display: none; }
                }
            </style>
        </head>
        <body>
            <h1>Quote Comparison Report</h1>
            <p>Generated: ${new Date().toLocaleString()}</p>
            <div class="summary-cards-grid">${summaryHTML}</div>
            ${tableHTML}
        </body>
        </html>
    `);
    
    printWindow.document.close();
    
    // Wait for content to load, then print
    setTimeout(() => {
        printWindow.print();
    }, 250);
    
    Helpers.showSuccess('Opening print dialog for PDF export');
}

/**
 * Handle accept quote from modal
 */
function handleAcceptQuoteFromModal(quote) {
    // Close modal first
    closeQuoteComparisonModal();
    
    // Then handle acceptance (reuse existing logic)
    if (typeof handleAcceptQuote === 'function') {
        // Find the quote in AppState or pass it directly
        handleAcceptQuote(quote);
    } else {
        Helpers.showSuccess(`Quote from ${quote.supplier_name} accepted`);
    }
}

/**
 * Close quote comparison modal
 */
function closeQuoteComparisonModal() {
    const modal = document.getElementById('quote-comparison-modal');
    if (modal) {
        modal.classList.add('hidden');
    }
    
    // Reset state
    modalQuotesState = {
        allQuotes: [],
        filteredQuotes: [],
        sortBy: 'unit_price_asc',
        filters: {
            search: '',
            bestPrice: false,
            fastestDelivery: false
        }
    };
}

// ==================== SETTINGS ====================
function openSettingsModal() {
    // Load current settings
    document.getElementById('api-url').value = Config.API_BASE_URL;
    document.getElementById('engineering-email').value = Config.ENGINEERING_EMAIL;
    document.getElementById('auto-classify').checked = 
        Config.getSetting(Config.STORAGE_KEYS.AUTO_CLASSIFY, true);
    document.getElementById('auto-create-folders').checked = 
        Config.getSetting(Config.STORAGE_KEYS.AUTO_CREATE_FOLDERS, true);
    
    // Load pin taskpane setting
    const isPinned = Config.getSetting('PIN_TASKPANE', false);
    document.getElementById('pin-taskpane').checked = isPinned;
    updatePinStatusMessage(isPinned);
    
    // Add change listener for pin checkbox
    document.getElementById('pin-taskpane').onchange = function() {
        updatePinStatusMessage(this.checked);
    };
    
    Helpers.showElement(document.getElementById('settings-modal'));
}

function updatePinStatusMessage(isPinned) {
    const statusEl = document.getElementById('pin-status');
    if (isPinned) {
        statusEl.innerHTML = '<span class="pin-enabled-message">✓ Add-in will stay open when you navigate between emails.</span>';
    } else {
        statusEl.textContent = 'Enable this to keep the add-in panel visible as you navigate between emails.';
    }
}

function closeSettingsModal() {
    Helpers.hideElement(document.getElementById('settings-modal'));
}

function saveSettings() {
    const isPinned = document.getElementById('pin-taskpane').checked;
    
    let apiUrl = document.getElementById('api-url').value.trim();
    
    // Prevent localhost URLs - always use production
    if (apiUrl.includes('localhost') || apiUrl.includes('127.0.0.1')) {
        showError('Cannot use localhost URL. Using production backend URL instead.');
        // Reset to production URL
        apiUrl = 'https://hexa-outlook-backend.onrender.com';
        document.getElementById('api-url').value = apiUrl;
    }
    
    const settings = {
        apiUrl: apiUrl,
        engineeringEmail: document.getElementById('engineering-email').value,
        autoClassify: document.getElementById('auto-classify').checked,
        autoCreateFolders: document.getElementById('auto-create-folders').checked
    };
    
    Config.saveSettings(settings);
    
    // Save pin setting separately
    Config.setSetting('PIN_TASKPANE', isPinned);
    
    // Apply pin behavior
    if (isPinned) {
        enablePinnedBehavior();
    }
    
    closeSettingsModal();
    Helpers.showSuccess('Settings saved' + (isPinned ? ' - Add-in is now pinned!' : ''));
}

/**
 * Enable pinned behavior - ensures the taskpane stays responsive
 * when navigating between emails
 */
function enablePinnedBehavior() {
    // The ItemChanged handler is already registered in initializeApp
    // This function ensures the add-in maintains its state
    console.log('Pinned behavior enabled - taskpane will stay open when changing emails');
    
    // Store the current state so it persists
    try {
        localStorage.setItem('procurement_addin_pinned', 'true');
    } catch (e) {
        console.log('Could not save pin state to localStorage');
    }
}

// Make functions available globally for inline event handlers
window.previewRFQ = previewRFQ;
window.createSingleDraft = createSingleDraft;
window.handleSupplierCheckboxChange = handleSupplierCheckboxChange;
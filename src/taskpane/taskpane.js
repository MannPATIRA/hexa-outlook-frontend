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
    emailContext: null
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
        
        // Clear the success flag but keep other state
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
    const oneHourAgo = Date.now() - (60 * 60 * 1000);
    if (state.timestamp && state.timestamp > oneHourAgo) {
        if (state.selectedPR) {
            AppState.selectedPR = state.selectedPR;
            console.log('Restored selected PR:', state.selectedPR);
        }
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
    
    // Show nav tabs and main content
    const navTabs = document.querySelector('.nav-tabs');
    const mainContent = document.getElementById('main-content');
    if (navTabs) navTabs.style.display = 'flex';
    if (mainContent) mainContent.style.display = 'block';
}

/**
 * Show a specific mode and hide the normal workflow
 */
function showMode(modeId) {
    console.log('showMode called with:', modeId);
    
    try {
        // Hide nav tabs and main content
        const navTabs = document.querySelector('.nav-tabs');
        const mainContent = document.getElementById('main-content');
        if (navTabs) navTabs.style.display = 'none';
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
            if (navTabs) navTabs.style.display = 'flex';
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
}

/**
 * Show Draft mode when user is viewing a draft email
 */
async function showDraftMode(context) {
    console.log('Showing Draft mode');
    showMode('draft-mode');
    AppState.currentMode = 'draft';
    
    // Load pending RFQ drafts
    await loadPendingDrafts();
}

/**
 * Load and display pending RFQ drafts
 */
async function loadPendingDrafts() {
    const listContainer = document.getElementById('pending-drafts-list');
    if (!listContainer) return;
    
    listContainer.innerHTML = '<p class="loading-text">Loading drafts...</p>';
    
    if (!AuthService.isSignedIn()) {
        listContainer.innerHTML = '<p class="loading-text">Please sign in to view drafts</p>';
        return;
    }
    
    try {
        // Get drafts from the Drafts folder that look like RFQs
        const drafts = await AuthService.graphRequest(
            `/me/mailFolders/Drafts/messages?$filter=startswith(subject,'RFQ for')&$select=id,subject,toRecipients,createdDateTime&$top=20&$orderby=createdDateTime desc`
        );
        
        if (!drafts.value || drafts.value.length === 0) {
            listContainer.innerHTML = '<p class="loading-text">No RFQ drafts found. Generate RFQs in the workflow first.</p>';
            return;
        }
        
        // Render the drafts
        listContainer.innerHTML = drafts.value.map(draft => {
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
        
        // Enable send button
        const sendBtn = document.getElementById('send-all-drafts-btn');
        if (sendBtn) sendBtn.disabled = false;
        
    } catch (error) {
        console.error('Error loading drafts:', error);
        listContainer.innerHTML = '<p class="loading-text">Error loading drafts. Please try again.</p>';
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
        
        const email = context.email;
        const originalRfq = context.originalRfq; // May be present if opened from sent RFQ
        
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
        
        // Extract and display the question
        const questionBox = document.getElementById('clarification-question-text');
        if (questionBox) {
            if (email.body?.content) {
                const bodyText = Helpers.stripHtml(email.body.content);
                const truncatedBody = bodyText.length > 500 ? bodyText.substring(0, 500) + '...' : bodyText;
                questionBox.textContent = truncatedBody;
            } else {
                questionBox.textContent = 'Email body not available';
            }
        }
        
        // Get suggested response from API (don't await - let it load in background)
        loadSuggestedResponse(email).catch(err => {
            console.error('Error loading suggested response:', err);
        });
        
    } catch (error) {
        console.error('Error in showClarificationMode:', error);
        Helpers.showError('Error displaying clarification: ' + error.message);
    }
}

/**
 * Load suggested response for a clarification email
 */
async function loadSuggestedResponse(email) {
    const loadingEl = document.getElementById('suggested-answer-loading');
    const contentEl = document.getElementById('suggested-answer-content');
    const textareaEl = document.getElementById('clarification-response-text');
    
    if (loadingEl) loadingEl.classList.remove('hidden');
    if (contentEl) contentEl.classList.add('hidden');
    
    try {
        // Extract question from body
        const bodyText = email.body?.content ? Helpers.stripHtml(email.body.content) : '';
        
        let suggestedResponse = '';
        
        // Try API first
        try {
            const result = await ApiClient.suggestResponse(
                email.id,
                email.id,
                bodyText.substring(0, 1000) // Limit question length
            );
            suggestedResponse = result.suggested_response || '';
            console.log('Suggested response from API received');
        } catch (apiError) {
            console.warn('API suggestion failed:', apiError.message);
            // Generate a basic template response
            suggestedResponse = generateFallbackResponse(email, bodyText);
        }
        
        if (loadingEl) loadingEl.classList.add('hidden');
        if (contentEl) contentEl.classList.remove('hidden');
        
        if (textareaEl) {
            textareaEl.value = suggestedResponse || 'Please compose your reply manually.';
        }
        
    } catch (error) {
        console.error('Error getting suggested response:', error);
        if (loadingEl) loadingEl.classList.add('hidden');
        if (contentEl) contentEl.classList.remove('hidden');
        if (textareaEl) {
            textareaEl.value = generateFallbackResponse(email, '');
        }
    }
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
 * Load and parse quote data from email
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
        
        // Try API first
        try {
            const result = await ApiClient.extractQuote(
                email.id,
                rfqId,
                supplierEmail,
                bodyContent
            );
            details = result.extracted_details || result || {};
            console.log('Quote extracted via API:', details);
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
    const statusEl = document.getElementById('draft-send-status');
    const progressText = document.getElementById('draft-progress-text');
    const progressFill = document.getElementById('draft-progress-fill');
    
    try {
        // Show progress UI
        if (sendBtn) sendBtn.disabled = true;
        if (statusEl) statusEl.classList.remove('hidden');
        
        // Get all RFQ drafts
        if (progressText) progressText.textContent = 'Finding RFQ drafts...';
        const draftsResponse = await AuthService.graphRequest(
            `/me/mailFolders/Drafts/messages?$filter=startswith(subject,'RFQ for')&$select=id,subject,toRecipients,body&$top=50`
        );
        
        const drafts = draftsResponse.value || [];
        
        if (drafts.length === 0) {
            Helpers.showError('No RFQ drafts found');
            if (sendBtn) sendBtn.disabled = false;
            if (statusEl) statusEl.classList.add('hidden');
            return;
        }
        
        console.log(`Found ${drafts.length} RFQ drafts to send`);
        
        // Get current email ID (the draft we might be viewing)
        const currentDraftId = Office.context.mailbox.item?.itemId;
        console.log('Currently viewing draft ID:', currentDraftId);
        
        // Separate current draft from others - we'll send it LAST
        const otherDrafts = drafts.filter(d => d.id !== currentDraftId);
        const currentDraft = drafts.find(d => d.id === currentDraftId);
        
        // Persist initial state
        persistState({
            sendingInProgress: true,
            totalDrafts: drafts.length,
            sentCount: 0,
            autoRepliesScheduled: 0
        });
        
        let sentCount = 0;
        let autoRepliesScheduled = 0;
        const totalDrafts = drafts.length;
        
        // STEP 1: Send ALL OTHER drafts first (not the current one)
        // For each, complete the ENTIRE workflow before moving to next
        for (const draft of otherDrafts) {
            try {
                const recipient = draft.toRecipients?.[0]?.emailAddress?.address || 'unknown';
                if (progressText) progressText.textContent = `Sending to ${recipient}... (${sentCount + 1}/${totalDrafts})`;
                if (progressFill) progressFill.style.width = `${((sentCount + 0.3) / totalDrafts) * 100}%`;
                
                // Send the draft and get the sent email details
                const sendResult = await sendDraftEmailWithFullWorkflow(draft);
                sentCount++;
                
                // Update UI
                const draftItem = document.querySelector(`[data-draft-id="${draft.id}"]`);
                if (draftItem) {
                    const statusBadge = draftItem.querySelector('.draft-item-status');
                    if (statusBadge) {
                        statusBadge.textContent = 'Sent';
                        statusBadge.classList.add('sent');
                    }
                }
                
                if (progressFill) progressFill.style.width = `${((sentCount) / totalDrafts) * 100}%`;
                
                if (sendResult.autoReplyScheduled) {
                    autoRepliesScheduled++;
                }
                
                // Update persisted state after each successful send
                persistState({ sentCount, autoRepliesScheduled });
                
                console.log(`✓ Sent ${sentCount}/${totalDrafts}: ${draft.subject}`);
                
            } catch (error) {
                console.error(`✗ Failed to send draft to ${draft.toRecipients?.[0]?.emailAddress?.address}:`, error);
            }
        }
        
        // STEP 2: If we sent all non-current drafts and there's no current draft, we're done
        if (!currentDraft) {
            persistState({ 
                sendingInProgress: false, 
                lastSendResult: 'success',
                sentCount,
                autoRepliesScheduled
            });
            if (progressFill) progressFill.style.width = '100%';
            if (progressText) progressText.textContent = `✓ Sent ${sentCount} RFQ(s)! Auto-replies: ${autoRepliesScheduled}`;
            Helpers.showSuccess(`Sent ${sentCount} RFQ(s) successfully! ${autoRepliesScheduled} auto-replies scheduled.`);
            return;
        }
        
        // STEP 3: Send the CURRENT draft last
        // After this, the add-in WILL close because we're viewing this draft
        if (progressText) progressText.textContent = `Sending final draft... Panel will close shortly.`;
        if (progressFill) progressFill.style.width = '95%';
        
        // Mark state as complete BEFORE sending current draft (because we won't get a chance after)
        persistState({ 
            sendingInProgress: false, 
            lastSendResult: 'success',
            sentCount: sentCount + 1, // Include the one we're about to send
            autoRepliesScheduled: autoRepliesScheduled + 1, // Assume it will work
            showSuccessOnReopen: true
        });
        
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
    } finally {
        if (sendBtn) sendBtn.disabled = false;
    }
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
                    delaySeconds: 30,
                    quantity: quantity
                });
                
                console.log(`✓ Auto-reply scheduled (will arrive in ~30 seconds)`);
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
    
    try {
        Helpers.showLoading('Forwarding to engineering...');
        
        const email = AppState.emailContext.email;
        const engineeringEmail = Config.getSetting(Config.STORAGE_KEYS.ENGINEERING_EMAIL, 'engineering@company.com');
        
        // Create forward draft
        const subject = `[Engineering Review] ${email.subject}`;
        const body = `
            <p>Please review the following technical clarification request:</p>
            <hr>
            <p><strong>Original Email:</strong></p>
            <p><strong>From:</strong> ${email.from?.emailAddress?.address}</p>
            <p><strong>Subject:</strong> ${email.subject}</p>
            <hr>
            ${email.body?.content || ''}
        `;
        
        await EmailOperations.createDraft(engineeringEmail, subject, body);
        
        // Move original email to Awaiting Engineer folder
        const materialMatch = email.subject?.match(/MAT-\d+/i);
        if (materialMatch) {
            const folderPath = `${materialMatch[0]}/${Config.FOLDERS.AWAITING_ENGINEER}`;
            try {
                await FolderManagement.moveEmailToFolder(email.id, folderPath);
            } catch (e) {
                console.error('Could not move email to folder:', e);
            }
        }
        
        Helpers.showSuccess('Forwarded to engineering team');
        
        // Go back to workflow
        showRFQWorkflowMode();
        
    } catch (error) {
        Helpers.showError('Failed to forward: ' + error.message);
    } finally {
        Helpers.hideLoading();
    }
}

/**
 * Handle replying to supplier with clarification response
 */
async function handleReplyToSupplier() {
    if (!AppState.emailContext?.email) {
        Helpers.showError('No email context');
        return;
    }
    
    const responseText = document.getElementById('clarification-response-text')?.value;
    if (!responseText || responseText.trim().length === 0) {
        Helpers.showError('Please enter a response');
        return;
    }
    
    try {
        Helpers.showLoading('Creating reply...');
        
        const email = AppState.emailContext.email;
        const supplierEmail = email.from?.emailAddress?.address;
        
        // Create reply
        const subject = email.subject.startsWith('RE:') ? email.subject : `RE: ${email.subject}`;
        const htmlBody = EmailOperations.formatTextAsHtml(responseText);
        
        await EmailOperations.createDraft(supplierEmail, subject, htmlBody);
        
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
        
        Helpers.showSuccess('Reply draft created');
        
        // Go back to workflow
        showRFQWorkflowMode();
        
    } catch (error) {
        Helpers.showError('Failed to create reply: ' + error.message);
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
    
        // Set up event listeners FIRST (so UI is responsive)
    setupEventListeners();
        setupModeEventListeners();
        
        // Initialize authentication - MUST await to ensure auth before context detection
        await initializeAuth();
    
        // Register for ItemChanged event
        registerItemChangedHandler();
        
        // Restore persisted state (shows success message if we were sending)
        // Note: This only shows messages, we ALWAYS detect context afterwards
        try {
            restorePersistedState();
        } catch (e) {
            console.error('Error restoring state:', e);
        }
        
        // ALWAYS detect email context and render appropriate UI
        // Even if we showed a success message, we need to detect what email we're on
        // CRITICAL: Wrap context detection in its own try-catch
        // so the add-in always shows something
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
    
    // Clarification mode buttons
    document.getElementById('send-to-engineer-btn')?.addEventListener('click', handleSendToEngineer);
    document.getElementById('reply-to-supplier-btn')?.addEventListener('click', handleReplyToSupplier);
    
    // Quote mode buttons
    document.getElementById('compare-quotes-btn')?.addEventListener('click', () => {
        // Switch to quote comparison tab
        showRFQWorkflowMode();
        const quoteTab = document.querySelector('[data-tab="quote-comparison"]');
        if (quoteTab) quoteTab.click();
    });
    document.getElementById('accept-quote-btn')?.addEventListener('click', handleAcceptQuote);
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
    
    // If in workflow mode, also update email processing tab info
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
    const userInfo = document.getElementById('user-info');
    const userName = document.getElementById('user-name');

    if (AuthService.isSignedIn()) {
        const user = AuthService.getUser();
        signInBtn?.classList.add('hidden');
        userInfo?.classList.remove('hidden');
        if (userName && user) {
            userName.textContent = user.name || user.email;
        }
    } else {
        signInBtn?.classList.remove('hidden');
        userInfo?.classList.add('hidden');
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
    document.querySelectorAll('.nav-tab').forEach(tab => {
        tab.addEventListener('click', handleTabClick);
    });

    // Refresh button
    document.getElementById('refresh-btn')?.addEventListener('click', handleRefresh);

    // Settings button
    document.getElementById('settings-btn')?.addEventListener('click', openSettingsModal);
    document.getElementById('close-settings')?.addEventListener('click', closeSettingsModal);
    document.getElementById('save-settings')?.addEventListener('click', saveSettings);

    // PR search
    document.getElementById('pr-search')?.addEventListener('input', 
        Helpers.debounce(handlePRSearch, 300));

    // Select all suppliers
    document.getElementById('select-all-suppliers')?.addEventListener('change', handleSelectAllSuppliers);

    // Generate RFQs button
    document.getElementById('generate-rfqs-btn')?.addEventListener('click', handleGenerateRFQs);

    // Send all RFQs
    document.getElementById('send-all-rfqs-btn')?.addEventListener('click', handleSendAllRFQs);

    // Email processing
    document.getElementById('classify-email-btn')?.addEventListener('click', handleClassifyEmail);
    document.getElementById('extract-quote-btn')?.addEventListener('click', handleExtractQuote);
    document.getElementById('send-response-btn')?.addEventListener('click', handleSendClarificationResponse);
    document.getElementById('forward-to-engineering-btn')?.addEventListener('click', handleForwardToEngineering);
    document.getElementById('process-engineer-response-btn')?.addEventListener('click', handleProcessEngineerResponse);
    document.getElementById('create-engineer-draft-btn')?.addEventListener('click', handleCreateEngineerDraft);

    // Quote comparison
    document.getElementById('rfq-select')?.addEventListener('change', handleRFQSelect);

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
    
    // Pin reminder banner
    document.getElementById('dismiss-pin-reminder')?.addEventListener('click', dismissPinReminder);
    
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
    const pinReminderBanner = document.getElementById('pin-reminder-banner');
    pinReminderBanner?.classList.add('hidden');
    localStorage.setItem('procurement_pin_reminder_dismissed', 'true');
    console.log('Pin reminder dismissed and saved');
}

// ==================== TAB NAVIGATION ====================
function handleTabClick(event) {
    const clickedTab = event.currentTarget;
    const tabName = clickedTab.dataset.tab;
    
    // Update tab buttons
    document.querySelectorAll('.nav-tab').forEach(tab => {
        tab.classList.remove('active');
    });
    clickedTab.classList.add('active');
    
    // Update tab content
    document.querySelectorAll('.tab-content').forEach(content => {
        content.classList.remove('active');
        Helpers.hideElement(content);
    });
    
    const targetContent = document.getElementById(`${tabName}-tab`);
    if (targetContent) {
        targetContent.classList.add('active');
        Helpers.showElement(targetContent);
    }
    
    // Load data for specific tabs
    if (tabName === 'email-processing') {
        loadCurrentEmailInfo();
    } else if (tabName === 'quote-comparison') {
        loadAvailableRFQs();
    }
}

// ==================== DATA LOADING ====================
async function loadInitialData() {
    // Check if we have persisted state with PRs
    const state = getPersistedState();
    
    if (state.prs && state.prs.length > 0) {
        // Restore PRs from state
        AppState.prs = state.prs;
        renderPRList(AppState.prs);
        
        // Also restore selected PR if any
        if (state.selectedPR) {
            AppState.selectedPR = state.selectedPR;
            // Re-select it in the UI
            setTimeout(() => {
                const prItem = document.querySelector(`[data-pr-id="${state.selectedPR.pr_id}"]`);
                if (prItem) prItem.classList.add('selected');
            }, 100);
        }
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
    
    // Persist the selection
    persistState({ 
        selectedPR: pr,
        prs: AppState.prs,
        currentStep: 'suppliers'
    });
    
    // Show PR details
    const detailsContainer = document.getElementById('pr-details');
    if (detailsContainer) {
        detailsContainer.innerHTML = `
            <p><strong>PR ID:</strong> ${Helpers.escapeHtml(pr.pr_id)}</p>
            <p><strong>Material:</strong> ${Helpers.escapeHtml(pr.material || 'N/A')}</p>
            <p><strong>Quantity:</strong> ${pr.quantities || 'N/A'} ${pr.unit || ''}</p>
            <p><strong>Description:</strong> ${Helpers.escapeHtml(pr.description || 'N/A')}</p>
        `;
    }
    Helpers.showElement(document.getElementById('selected-pr-info'));
    
    // Enable supplier step and load suppliers
    Helpers.enableStep(document.getElementById('step-select-suppliers'));
    
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
    AppState.selectedSuppliers = [];
    
    if (countEl) {
        countEl.textContent = `${suppliers.length} supplier${suppliers.length !== 1 ? 's' : ''} found`;
    }
    
    if (suppliers.length === 0) {
        container.innerHTML = '<p class="placeholder-text">No matching suppliers found</p>';
        return;
    }
    
    suppliers.forEach(supplier => {
        const scoreClass = Helpers.getMatchScoreClass(supplier.match_score);
        
        const item = Helpers.createElement('div', {
            className: 'list-item',
            dataset: { supplierId: supplier.supplier_id }
        }, `
            <input type="checkbox" class="supplier-checkbox" 
                   data-supplier-id="${supplier.supplier_id}"
                   onchange="handleSupplierCheckboxChange(this)">
            <div class="supplier-info">
                <div class="list-item-title">
                    ${Helpers.escapeHtml(supplier.name)}
                    <span class="match-score ${scoreClass}">${supplier.match_score}/10</span>
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
    
    updateGenerateRFQsButton();
}

function handleSupplierCheckboxChange(checkbox) {
    const supplierId = checkbox.dataset.supplierId;
    
    if (checkbox.checked) {
        if (!AppState.selectedSuppliers.includes(supplierId)) {
            AppState.selectedSuppliers.push(supplierId);
        }
    } else {
        AppState.selectedSuppliers = AppState.selectedSuppliers.filter(id => id !== supplierId);
    }
    
    // Update select all checkbox
    const allCheckboxes = document.querySelectorAll('.supplier-checkbox');
    const selectAllCheckbox = document.getElementById('select-all-suppliers');
    if (selectAllCheckbox) {
        selectAllCheckbox.checked = AppState.selectedSuppliers.length === allCheckboxes.length;
    }
    
    updateGenerateRFQsButton();
}

function handleSelectAllSuppliers(event) {
    const isChecked = event.target.checked;
    const checkboxes = document.querySelectorAll('.supplier-checkbox');
    
    AppState.selectedSuppliers = [];
    
    checkboxes.forEach(cb => {
        cb.checked = isChecked;
        if (isChecked) {
            AppState.selectedSuppliers.push(cb.dataset.supplierId);
        }
    });
    
    updateGenerateRFQsButton();
}

function updateGenerateRFQsButton() {
    const btn = document.getElementById('generate-rfqs-btn');
    if (btn) {
        btn.disabled = AppState.selectedSuppliers.length === 0;
        btn.textContent = AppState.selectedSuppliers.length > 0
            ? `Generate RFQs (${AppState.selectedSuppliers.length})`
            : 'Generate RFQs';
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
            Helpers.showLoading('Saving drafts...');
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
                    
                    // Save draft
                    const draft = await EmailOperations.saveDraft({
                        to: [rfq.supplier_email],
                        subject: rfq.subject || '',
                        body: htmlBody,
                        cc: []
                    });
                    
                    // Store draft ID in RFQ object
                    rfq.draftId = draft.id;
                } catch (error) {
                    console.error(`Failed to save draft for ${rfq.supplier_name}:`, error);
                    // Continue with other RFQs even if one fails
                }
            }
        }
        
        AppState.rfqs = rfqs;
        
        // Enable review step
        Helpers.enableStep(document.getElementById('step-review-rfqs'));
        
        // Render RFQ cards
        renderRFQCards(rfqs);
        
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
    
    // Enable action buttons
    document.getElementById('send-all-rfqs-btn').disabled = false;
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
            htmlBody
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

async function handleSendAllRFQs() {
    if (!AuthService.isSignedIn()) {
        Helpers.showError('Please sign in to send emails');
        return;
    }

    if (!AppState.selectedPR || AppState.rfqs.length === 0) {
        Helpers.showError('No RFQs to send');
        return;
    }

    try {
        Helpers.showLoading('Sending RFQs...');

        // Extract material code from selected PR
        const materialCode = Helpers.extractMaterialCode(AppState.selectedPR);

        // Create folder structure before sending
        try {
            await FolderManagement.initializeMaterialFolders(materialCode);
        } catch (error) {
            console.error('Failed to create folder structure:', error);
            // Continue sending even if folder creation fails
        }

        let successCount = 0;
        let failCount = 0;

        // Send each RFQ
        for (const rfq of AppState.rfqs) {
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

                // Validate email address
                if (!rfq.supplier_email || !rfq.supplier_email.includes('@')) {
                    throw new Error(`Invalid email address: ${rfq.supplier_email}`);
                }

                // Send email with materialCode for folder organization
                console.log(`Sending RFQ to ${rfq.supplier_name} (${rfq.supplier_email})`);
                const sendResult = await EmailOperations.sendEmail({
                    to: [rfq.supplier_email],
                    subject: rfq.subject || '',
                    body: htmlBody,
                    cc: [],
                    materialCode: materialCode
                });

                // Delete the draft if it exists (draft was auto-created when RFQs were generated)
                if (rfq.draftId) {
                    try {
                        console.log(`Deleting draft ${rfq.draftId} for ${rfq.supplier_name}`);
                        await EmailOperations.deleteDraft(rfq.draftId);
                        // Remove draftId from RFQ object since it's been deleted
                        delete rfq.draftId;
                    } catch (error) {
                        console.error(`Failed to delete draft for ${rfq.supplier_name}:`, error);
                        // Continue even if draft deletion fails
                    }
                }

                // Schedule auto-reply for demo/testing purposes
                if (sendResult.internetMessageId) {
                    try {
                        const userEmail = Office.context.mailbox.userProfile?.emailAddress;
                        const materialName = AppState.selectedPR?.material || 
                                           AppState.selectedPR?.description || 
                                           materialCode;
                        const quantity = AppState.selectedPR?.quantity || 100;

                        if (userEmail) {
                            console.log(`Scheduling auto-reply for RFQ to ${rfq.supplier_name}...`);
                            const replyResult = await ApiClient.scheduleAutoReply({
                                toEmail: userEmail,
                                subject: rfq.subject || '',
                                internetMessageId: sendResult.internetMessageId,
                                material: materialName,
                                replyType: 'random',
                                delaySeconds: 30,
                                quantity: quantity
                            });
                            console.log(`✓ Auto-reply scheduled for ${rfq.supplier_name}:`, replyResult);
                        } else {
                            console.warn('Could not get user email for auto-reply scheduling');
                        }
                    } catch (replyError) {
                        // Non-critical: don't fail the send if auto-reply scheduling fails
                        console.error(`Failed to schedule auto-reply for ${rfq.supplier_name}:`, replyError);
                    }
                } else {
                    console.warn(`No internetMessageId available for ${rfq.supplier_name} - auto-reply not scheduled`);
                }

                console.log(`Successfully sent RFQ to ${rfq.supplier_name}`);
                successCount++;
            } catch (error) {
                console.error(`Failed to send RFQ to ${rfq.supplier_name} (${rfq.supplier_email}):`, error);
                console.error('Error details:', error.message, error.stack);
                failCount++;
                // Continue with other RFQs even if one fails
            }
        }

        // Update RFQ statuses and re-render (drafts have been deleted, so buttons will update)
        AppState.rfqs.forEach(rfq => {
            rfq.status = 'sent';
            // Drafts have been deleted, so remove draftId
            if (rfq.draftId) {
                delete rfq.draftId;
            }
        });
        renderRFQCards(AppState.rfqs);

        if (failCount === 0) {
            Helpers.showSuccess(`All ${successCount} RFQ(s) sent successfully. Replies will arrive in ~30 seconds.`);
        } else {
            Helpers.showError(`${successCount} sent, ${failCount} failed`);
        }
    } catch (error) {
        Helpers.showError('Failed to send RFQs: ' + error.message);
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
        container.innerHTML = '<p class="placeholder-text">No quotes received yet</p>';
        Helpers.hideElement(document.getElementById('quote-summary'));
        return;
    }
    
    quotes.forEach(quote => {
        const card = Helpers.createElement('div', {
            className: 'quote-card'
        }, `
            <div class="quote-card-header">
                <h4>${Helpers.escapeHtml(quote.supplier_name)}</h4>
            </div>
            <div class="quote-card-body">
                <div class="quote-field">
                    <label>Price</label>
                    <span class="value price">${Helpers.formatCurrency(quote.price, quote.currency)}</span>
                </div>
                <div class="quote-field">
                    <label>Delivery</label>
                    <span class="value">${Helpers.escapeHtml(quote.delivery_time || 'N/A')}</span>
                </div>
                <div class="quote-field">
                    <label>Quote Date</label>
                    <span class="value">${Helpers.formatDate(quote.quote_date)}</span>
                </div>
                <div class="quote-field">
                    <label>Status</label>
                    <span class="value">${quote.status}</span>
                </div>
            </div>
        `);
        
        container.appendChild(card);
    });
    
    // Show summary
    if (quotes.length > 1) {
        const lowest = Helpers.sortBy(quotes, 'price')[0];
        const summaryContainer = document.getElementById('summary-content');
        if (summaryContainer) {
            summaryContainer.innerHTML = `
                <p><strong>Lowest Price:</strong> ${Helpers.formatCurrency(lowest.price, lowest.currency)} 
                   from ${Helpers.escapeHtml(lowest.supplier_name)}</p>
                <p><strong>Total Quotes:</strong> ${quotes.length}</p>
            `;
        }
        Helpers.showElement(document.getElementById('quote-summary'));
    }
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
    
    const settings = {
        apiUrl: document.getElementById('api-url').value,
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
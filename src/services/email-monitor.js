/**
 * Get RFQ mapping from email (uses global RFQMapping if available)
 * @param {Object} email - Email object with inReplyTo/references
 * @returns {Object|null} Mapping object or null if not found
 */
function getRFQMappingFromEmail(email) {
    if (typeof window !== 'undefined' && window.RFQMapping) {
        return window.RFQMapping.getFromEmail(email);
    }
    return null;
}

/**
 * Email Monitoring Service
 * Automatically monitors for replies to emails in Sent RFQs folder
 * Classifies and organizes them automatically
 */
const EmailMonitor = {
    // Track processed emails to avoid duplicates
    processedEmails: new Set(),
    
    // Monitoring interval (in milliseconds)
    monitoringInterval: 3000, // 3 seconds
    
    // Interval ID for cleanup
    intervalId: null,
    
    // Whether monitoring is active
    isMonitoring: false,
    
    // Logging prefix for easy identification
    LOG_PREFIX: '[EmailMonitor]',

    /**
     * Log helper with prefix
     */
    log(...args) {
        console.log(this.LOG_PREFIX, ...args);
    },

    /**
     * Error log helper with prefix
     */
    logError(...args) {
        console.error(this.LOG_PREFIX, '❌', ...args);
    },

    /**
     * Success log helper with prefix
     */
    logSuccess(...args) {
        console.log(this.LOG_PREFIX, '✓', ...args);
    },

    /**
     * Start monitoring for replies to Sent RFQs
     */
    async startMonitoring() {
        if (this.isMonitoring) {
            this.log('Email monitoring already active');
            return;
        }

        if (!AuthService.isSignedIn()) {
            this.log('Not signed in - email monitoring will start after sign-in');
            return;
        }

        this.log('========================================');
        this.log('Starting automatic email monitoring...');
        this.log(`Check interval: ${this.monitoringInterval / 1000} seconds`);
        this.log('========================================');
        this.isMonitoring = true;

        // Initial check
        await this.checkForReplies();

        // Set up periodic checking
        this.intervalId = setInterval(async () => {
            if (AuthService.isSignedIn()) {
                await this.checkForReplies();
            }
        }, this.monitoringInterval);
    },

    /**
     * Stop monitoring
     */
    stopMonitoring() {
        if (this.intervalId) {
            clearInterval(this.intervalId);
            this.intervalId = null;
        }
        this.isMonitoring = false;
        this.log('Email monitoring stopped');
    },

    /**
     * Manually trigger a check for replies (for testing/debugging)
     * Call this from browser console: EmailMonitor.manualCheck()
     */
    async manualCheck() {
        this.log('========================================');
        this.log('MANUAL CHECK TRIGGERED');
        this.log('========================================');
        
        // Clear processed emails to recheck everything
        this.processedEmails.clear();
        this.log('Cleared processed emails cache');
        
        await this.checkForReplies();
    },

    /**
     * Check for new replies to Sent RFQs
     */
    async checkForReplies() {
        try {
            if (!AuthService.isSignedIn()) {
                this.log('Not signed in, skipping check');
                return;
            }

            this.log('========================================');
            this.log('CHECKING FOR REPLIES TO SENT RFQs');
            this.log(`Time: ${new Date().toISOString()}`);
            this.log('========================================');

            // Get recent emails from Inbox
            const recentEmails = await this.getRecentInboxEmails();
            this.log(`Found ${recentEmails.length} recent email(s) in inbox`);
            
            if (recentEmails.length === 0) {
                this.log('No emails to check');
                return;
            }
            
            // Log all emails for debugging
            this.log('Emails in inbox:');
            recentEmails.forEach((e, i) => {
                this.log(`  ${i + 1}. "${e.subject}" from ${e.from?.emailAddress?.address || 'unknown'}`);
            });
            
            let processedCount = 0;
            let skippedCount = 0;
            let detectedCount = 0;
            
            for (const email of recentEmails) {
                // Skip if already processed
                if (this.processedEmails.has(email.id)) {
                    skippedCount++;
                    continue;
                }

                this.log('----------------------------------------');
                this.log(`Checking email: "${email.subject}"`);
                this.log(`  From: ${email.from?.emailAddress?.address || 'unknown'}`);
                this.log(`  ID: ${email.id}`);

                // CRITICAL: Check if email is from Microsoft Outlook and delete immediately
                if (EmailOperations.isFromMicrosoftOutlook(email)) {
                    this.log('⚠️  DETECTED Microsoft Outlook email - DELETING immediately');
                    try {
                        await EmailOperations.deleteEmail(email.id);
                        this.logSuccess(`✓ Deleted Microsoft Outlook email: ${email.id}`);
                        this.processedEmails.add(email.id);
                        skippedCount++;
                        continue; // Skip all further processing
                    } catch (deleteError) {
                        this.logError('Failed to delete Microsoft Outlook email:', deleteError.message);
                        // Continue to mark as processed so we don't keep trying
                        this.processedEmails.add(email.id);
                        skippedCount++;
                        continue;
                    }
                }

                // Check if this is a reply to an email with SENT RFQ category
                const isReplyToSentRfq = await this.isReplyToSentRfq(email);
                
                if (isReplyToSentRfq) {
                    detectedCount++;
                    this.logSuccess(`>>> DETECTED as reply to Sent RFQ!`);
                    this.log(`    Material Code: ${isReplyToSentRfq.materialCode || 'unknown'}`);
                    
                    try {
                        await this.processReply(email, isReplyToSentRfq);
                        this.processedEmails.add(email.id);
                        processedCount++;
                    } catch (processError) {
                        this.logError('Failed to process reply:', processError.message);
                    }
                } else {
                    this.log(`  Not an RFQ reply, skipping`);
                }
            }
            
            this.log('========================================');
            this.log('CHECK COMPLETE');
            this.log(`  Detected: ${detectedCount} RFQ replies`);
            this.log(`  Processed: ${processedCount} emails`);
            this.log(`  Skipped: ${skippedCount} (already processed)`);
            this.log('========================================');
        } catch (error) {
            this.logError('Error checking for replies:', error.message);
            this.logError('Stack:', error.stack);
        }
    },

    /**
     * Get recent emails from Inbox (both read and unread for reliability)
     */
    async getRecentInboxEmails() {
        try {
            // Get both read and unread recent emails to ensure we catch replies
            // Removed isRead filter to catch emails even if user viewed them
            const endpoint = '/me/mailFolders/inbox/messages?' +
                '$select=id,subject,from,receivedDateTime,conversationId,internetMessageId,bodyPreview,isRead' +
                '&$top=20' +
                '&$orderby=receivedDateTime desc';
            
            const response = await AuthService.graphRequest(endpoint);
            const emails = response.value || [];
            this.log(`Fetched ${emails.length} recent emails from inbox`);
            return emails;
        } catch (error) {
            this.logError('Error getting recent inbox emails:', error.message);
            return [];
        }
    },

    /**
     * Check if an email is a reply to an email with SENT RFQ category
     * Returns the parent email info if found, null otherwise
     * 
     * Detection methods (in order):
     * 1. Subject-based: Check if subject contains "Re:" and "MAT-XXXXX" pattern
     * 2. Conversation-based: Look for emails in same conversation with SENT RFQ category
     * 3. Folder-based: Check if email is already in a Sent RFQs folder
     */
    async isReplyToSentRfq(email) {
        try {
            const subject = email.subject || '';
            this.log(`  Analyzing email subject: "${subject}"`);
            
            // ===========================================
            // METHOD 1: Subject-based detection (PRIMARY)
            // This is the most reliable method as it doesn't depend on Graph API folder search
            // ===========================================
            
            // Check if this looks like a reply (starts with Re:, RE:, Fwd:, etc.)
            const isReply = /^(re:|fw:|fwd:)/i.test(subject.trim());
            
            // Check if subject contains RFQ pattern (e.g., "RFQ for MAT-12345")
            const containsRfqPattern = /rfq/i.test(subject);
            
            // Extract material code from subject
            const materialMatch = subject.match(/MAT-\d+/i);
            
            this.log(`  Subject analysis: isReply=${isReply}, containsRfqPattern=${containsRfqPattern}, materialCode=${materialMatch ? materialMatch[0] : 'none'}`);
            
            if (materialMatch && (isReply || containsRfqPattern)) {
                const materialCode = materialMatch[0].toUpperCase();
                this.logSuccess(`Detected RFQ reply via subject pattern! Material: ${materialCode}`);
                return {
                    parentEmailId: null,
                    parentSubject: subject,
                    materialCode: materialCode
                };
            }
            
            // ===========================================
            // METHOD 2: Conversation-based detection (FALLBACK)
            // Look for emails in the same conversation with SENT RFQ category
            // ===========================================
            this.log(`  Subject detection failed, trying conversation lookup...`);
            
            // Get full email details including conversation ID
            // Note: 'inReplyTo' is not available in Microsoft Graph, removed to avoid errors
            const fullEmail = await AuthService.graphRequest(
                `/me/messages/${email.id}?$select=id,subject,conversationId,internetMessageId`
            );

            if (!fullEmail.conversationId) {
                this.log(`  Email ${email.id} has no conversationId`);
            } else {
                this.log(`  Checking conversation: ${fullEmail.conversationId}`);

                // Escape special characters in conversationId for OData filter
                const escapedConversationId = fullEmail.conversationId
                    .replace(/'/g, "''")  // Escape single quotes
                    .replace(/\\/g, '\\\\'); // Escape backslashes

                try {
                    // Get all emails in the conversation
                    // Note: Personal Outlook accounts don't support $filter + $orderby together
                    // so we fetch without orderby and sort in JavaScript
                    const conversationEmails = await AuthService.graphRequest(
                        `/me/messages?$filter=conversationId eq '${escapedConversationId}'` +
                        `&$select=id,subject,categories,receivedDateTime` +
                        `&$top=50`
                    );
                    
                    // Sort by receivedDateTime in JavaScript
                    if (conversationEmails.value) {
                        conversationEmails.value.sort((a, b) => 
                            new Date(a.receivedDateTime) - new Date(b.receivedDateTime)
                        );
                    }

                    if (conversationEmails.value && conversationEmails.value.length > 0) {
                        this.log(`  Found ${conversationEmails.value.length} email(s) in conversation`);

                        // Find the original email in the conversation that has SENT RFQ category
                        for (const convEmail of conversationEmails.value) {
                            if (convEmail.id === email.id) {
                                continue; // Skip the current email
                            }

                            // Log categories for debugging
                            if (convEmail.categories && convEmail.categories.length > 0) {
                                this.log(`  Email ${convEmail.id} has categories: ${convEmail.categories.join(', ')}`);
                            }

                            // Check if this email has SENT RFQ category
                            if (convEmail.categories && 
                                convEmail.categories.some(cat => 
                                    cat.toLowerCase().includes('sent rfq')
                                )) {
                                // Found the parent email with SENT RFQ category
                                this.logSuccess(`Found parent email with SENT RFQ category: ${convEmail.id}`);
                                
                                // Extract material code from the folder path or subject
                                let materialCode = await this.extractMaterialCodeFromEmail(convEmail.id);
                                if (!materialCode) {
                                    // Try to extract from parent subject
                                    const parentMatch = convEmail.subject.match(/MAT-\d+/i);
                                    if (parentMatch) {
                                        materialCode = parentMatch[0].toUpperCase();
                                    }
                                }
                                this.log(`  Extracted material code: ${materialCode || 'none'}`);
                                
                                return {
                                    parentEmailId: convEmail.id,
                                    parentSubject: convEmail.subject,
                                    materialCode: materialCode
                                };
                            }
                        }
                    } else {
                        this.log(`  No emails found in conversation`);
                    }
                } catch (convError) {
                    this.logError('Conversation lookup failed:', convError.message);
                }
            }

            // ===========================================
            // METHOD 3: Folder-based detection (LAST RESORT)
            // Check if email is already in a Sent RFQs folder
            // ===========================================
            this.log(`  Conversation detection failed, checking folder...`);
            
            try {
                const emailWithFolder = await AuthService.graphRequest(
                    `/me/messages/${email.id}?$select=id,parentFolderId`
                );

                if (emailWithFolder.parentFolderId) {
                    const folderInfo = await this.getFolderInfo(emailWithFolder.parentFolderId);
                    if (folderInfo && folderInfo.displayName === 'Sent RFQs') {
                        this.logSuccess(`Email is in Sent RFQs folder`);
                        // Extract material code from parent folder
                        const materialCode = await this.extractMaterialCodeFromParentFolder(emailWithFolder.parentFolderId);
                        return {
                            parentEmailId: null,
                            parentSubject: email.subject,
                            materialCode: materialCode
                        };
                    }
                }
            } catch (folderError) {
                this.logError('Folder lookup failed:', folderError.message);
            }

            this.log(`  Email is NOT a reply to a Sent RFQ (all detection methods failed)`);
            return null;
        } catch (error) {
            this.logError('Error checking if reply to Sent RFQ:', error.message);
            this.logError('Stack:', error.stack);
            return null;
        }
    },

    /**
     * Get folder information by ID
     */
    async getFolderInfo(folderId) {
        try {
            return await AuthService.graphRequest(`/me/mailFolders/${folderId}?$select=id,displayName,parentFolderId`);
        } catch (error) {
            console.error('Error getting folder info:', error);
            return null;
        }
    },

    /**
     * Extract material code from email by checking its folder path
     */
    async extractMaterialCodeFromEmail(emailId) {
        try {
            const email = await AuthService.graphRequest(
                `/me/messages/${emailId}?$select=id,parentFolderId`
            );

            if (!email.parentFolderId) {
                return null;
            }

            return await this.extractMaterialCodeFromParentFolder(email.parentFolderId);
        } catch (error) {
            console.error('Error extracting material code from email:', error);
            return null;
        }
    },

    /**
     * Extract material code from folder path (e.g., "MAT-12345/Sent RFQs" -> "MAT-12345")
     */
    async extractMaterialCodeFromParentFolder(folderId) {
        try {
            let currentFolderId = folderId;
            const path = [];

            // Walk up the folder tree
            while (currentFolderId) {
                const folder = await this.getFolderInfo(currentFolderId);
                if (!folder) break;

                path.unshift(folder.displayName);
                
                // Check if this folder name looks like a material code (e.g., MAT-12345)
                if (folder.displayName && /^MAT-\d+$/i.test(folder.displayName)) {
                    return folder.displayName;
                }

                if (!folder.parentFolderId || folder.parentFolderId === 'inbox') {
                    break;
                }
                currentFolderId = folder.parentFolderId;
            }

            // If we didn't find a material code in the path, try to extract from subject
            // This is a fallback
            this.log(`  Could not find material code in folder path: ${path.join('/')}`);
            return null;
        } catch (error) {
            this.logError('Error extracting material code from folder:', error.message);
            return null;
        }
    },

    /**
     * Process a reply email: classify, categorize, and move to appropriate folder
     */
    async processReply(email, replyInfo) {
        this.log('========================================');
        this.log(`PROCESSING REPLY: ${email.id}`);
        this.log(`Subject: ${email.subject}`);
        this.log('========================================');
        
        try {
            // Step 0: CRITICAL - Check if email is from Microsoft Outlook and delete immediately
            this.log('Step 0: Checking for Microsoft Outlook emails...');
            if (EmailOperations.isFromMicrosoftOutlook(email)) {
                this.log('⚠️  DETECTED Microsoft Outlook email - DELETING immediately');
                try {
                    await EmailOperations.deleteEmail(email.id);
                    this.logSuccess(`✓ Deleted Microsoft Outlook email: ${email.id}`);
                    return; // Stop all processing
                } catch (deleteError) {
                    this.logError('Failed to delete Microsoft Outlook email:', deleteError.message);
                    throw deleteError; // Fail the processing
                }
            }

            // Step 1: Get full email details with body content
            this.log('Step 1: Fetching full email details...');
            const fullEmail = await AuthService.graphRequest(
                `/me/messages/${email.id}?$select=id,subject,from,body,receivedDateTime,conversationId`
            );
            this.logSuccess(`Got email from: ${fullEmail.from?.emailAddress?.address || 'unknown'}`);
            
            // Double-check after getting full email (in case from field wasn't complete)
            if (EmailOperations.isFromMicrosoftOutlook(fullEmail)) {
                this.log('⚠️  DETECTED Microsoft Outlook email (after full fetch) - DELETING immediately');
                try {
                    await EmailOperations.deleteEmail(email.id);
                    this.logSuccess(`✓ Deleted Microsoft Outlook email: ${email.id}`);
                    return; // Stop all processing
                } catch (deleteError) {
                    this.logError('Failed to delete Microsoft Outlook email:', deleteError.message);
                    throw deleteError; // Fail the processing
                }
            }
            
            // Step 2: Get email chain for classification
            this.log('Step 2: Building email chain for classification...');
            const emailChain = await this.getEmailChain(email.id);
            this.log(`  Email chain has ${emailChain.length} message(s)`);
            
            // Extract RFQ ID from subject if possible
            const rfqId = EmailOperations.extractRfqId(fullEmail.subject);
            this.log(`  Extracted RFQ ID: ${rfqId || 'none'}`);

            // Get email body content (prefer HTML, fallback to text)
            let emailBody = '';
            if (fullEmail.body) {
                emailBody = fullEmail.body.content || '';
            }
            this.log(`  Email body length: ${emailBody.length} chars`);

            // Step 3: Classify the email
            this.log('Step 3: Calling classification API...');
            
            // Get supplier info from RFQ mapping
            let rfqIdFromMapping = null;
            let supplierId = null;
            let supplierName = null;
            
            // Try to get mapping from email's inReplyTo or references
            const rfqMapping = getRFQMappingFromEmail(fullEmail);
            if (rfqMapping) {
                rfqIdFromMapping = rfqMapping.rfq_id;
                supplierId = rfqMapping.supplier_id;
                supplierName = rfqMapping.supplier_name;
                this.log(`  Found RFQ mapping: RFQ ${rfqIdFromMapping}, Supplier ${supplierName} (${supplierId})`);
            } else {
                // Fallback: Extract RFQ ID from subject
                rfqIdFromMapping = rfqId || EmailOperations.extractRfqId(fullEmail.subject);
                this.log(`  No RFQ mapping found, extracted RFQ ID from subject: ${rfqIdFromMapping || 'none'}`);
                
                // Fallback: Use sender email as supplier_id
                const senderEmail = fullEmail.from?.emailAddress?.address || '';
                supplierId = senderEmail || 'unknown';
                this.log(`  Using sender email as supplier_id: ${supplierId}`);
            }
            
            // Use rfqId from mapping if available, otherwise use extracted one
            const finalRfqId = rfqIdFromMapping || rfqId;
            
            const senderEmail = fullEmail.from?.emailAddress?.address || '';
            const classificationPayload = {
                subject: fullEmail.subject || '',
                body: emailBody,
                from_email: senderEmail,
                date: fullEmail.receivedDateTime || new Date().toISOString(),
                in_reply_to: finalRfqId
            };
            this.log(`  Classification request:`, JSON.stringify(classificationPayload, null, 2).substring(0, 500));
            this.log(`  Supplier ID: ${supplierId}, Supplier Name: ${supplierName || 'N/A'}`);
            
            let classification;
            try {
                classification = await ApiClient.classifyEmail(
                    emailChain,
                    classificationPayload,
                    finalRfqId,
                    supplierId  // Use supplier_id from mapping
                );
                this.logSuccess(`Classification result: ${classification.classification} (confidence: ${classification.confidence})`);
                
                // Store the backend email_id for future API calls
                // This mapping is used when user opens the email later
                if (classification.email_id) {
                    this.log(`  Storing backend email_id: ${classification.email_id}`);
                    // Use the storeEmailId function from taskpane.js if available
                    if (typeof storeEmailId === 'function') {
                        storeEmailId(email.id, classification.email_id);
                    } else {
                        // Fallback: store directly in localStorage
                        try {
                            const mapping = JSON.parse(localStorage.getItem('procurement_email_id_mapping') || '{}');
                            mapping[email.id] = classification.email_id;
                            localStorage.setItem('procurement_email_id_mapping', JSON.stringify(mapping));
                            this.log(`  Backend email_id stored in localStorage`);
                        } catch (storageError) {
                            this.logError('Failed to store email_id:', storageError.message);
                        }
                    }
                } else {
                    this.log(`  Warning: No email_id returned from classification API`);
                }
            } catch (classifyError) {
                this.logError('Classification API failed:', classifyError.message);
                this.logError('Full error:', classifyError);
                throw classifyError;
            }

            // Step 4: Determine material code
            this.log('Step 4: Determining material code...');
            let materialCode = replyInfo.materialCode;
            if (!materialCode) {
                // Try to extract from subject (e.g., "RFQ for MAT-12345")
                const match = fullEmail.subject.match(/MAT-\d+/i);
                if (match) {
                    materialCode = match[0];
                    this.log(`  Extracted from subject: ${materialCode}`);
                }
            } else {
                this.log(`  Using from replyInfo: ${materialCode}`);
            }

            if (!materialCode) {
                this.logError('Could not determine material code for email, skipping folder organization');
                return;
            }

            // Step 5: Determine target folder based on classification
            this.log('Step 5: Determining target folder...');
            const targetFolder = FolderManagement.getFolderForClassification(
                materialCode,
                classification.classification,
                classification.sub_classification || null
            );
            this.log(`  Target folder: ${targetFolder}`);

            // Step 6: Ensure folders exist
            this.log('Step 6: Initializing folder structure...');
            try {
                await FolderManagement.initializeMaterialFolders(materialCode);
                this.logSuccess('Folder structure ready');
            } catch (folderError) {
                this.logError('Failed to initialize folders:', folderError.message);
                throw folderError;
            }

            // Step 7: Move email to appropriate folder
            this.log('Step 7: Moving email to folder...');
            try {
                await FolderManagement.moveEmailToFolder(email.id, targetFolder);
                this.logSuccess(`Moved email to ${targetFolder}`);
            } catch (moveError) {
                this.logError('Failed to move email:', moveError.message);
                throw moveError;
            }

            // Step 8: Apply appropriate category based on classification
            this.log('Step 8: Applying category tag...');
            // Colors: Preset0=Red, Preset1=Orange, Preset2=Brown, Preset3=Yellow, 
            //         Preset4=Green, Preset5=Teal, Preset6=Blue, Preset7=Purple
            let categoryName = null;
            let categoryColor = 'Preset6'; // Default to blue
            
            switch (classification.classification) {
                case 'quote':
                    categoryName = 'QUOTE';
                    categoryColor = 'Preset4'; // Green
                    break;
                case 'clarification_request':
                    categoryName = 'CLARIFICATION';
                    categoryColor = 'Preset3'; // Yellow
                    break;
                case 'engineer_response':
                    categoryName = 'ENGINEER RESPONSE';
                    categoryColor = 'Preset6'; // Blue
                    break;
                default:
                    this.log(`  Unknown classification: ${classification.classification}, no category applied`);
            }

            if (categoryName) {
                try {
                    await EmailOperations.applyCategoryToEmail(email.id, categoryName, categoryColor);
                    this.logSuccess(`Applied category "${categoryName}" (${categoryColor}) to email`);
                } catch (categoryError) {
                    this.logError('Failed to apply category (non-critical):', categoryError.message);
                }
            }

            // Step 9: Mark email as read
            this.log('Step 9: Marking email as read...');
            try {
                await EmailOperations.markAsRead(email.id, true);
                this.logSuccess('Email marked as read');
            } catch (readError) {
                this.logError('Failed to mark email as read (non-critical):', readError.message);
            }

            this.log('========================================');
            this.logSuccess(`COMPLETED processing reply email ${email.id}`);
            this.log(`  Classification: ${classification.classification}`);
            this.log(`  Moved to: ${targetFolder}`);
            this.log(`  Category: ${categoryName || 'none'}`);
            this.log('========================================');
        } catch (error) {
            this.logError('========================================');
            this.logError('FAILED to process reply email:', email.id);
            this.logError('Error:', error.message);
            this.logError('Stack:', error.stack);
            this.logError('========================================');
            // Don't add to processed list if processing failed, so we can retry
        }
    },

    /**
     * Get email chain for classification
     */
    async getEmailChain(emailId) {
        try {
            const email = await EmailOperations.getEmailById(emailId);
            
            // Get conversation emails
            if (email.conversationId) {
                const conversationEmails = await AuthService.graphRequest(
                    `/me/messages?$filter=conversationId eq '${email.conversationId}'` +
                    `&$select=id,subject,from,body,receivedDateTime` +
                    `&$orderby=receivedDateTime asc`
                );

                if (conversationEmails.value && conversationEmails.value.length > 0) {
                    return conversationEmails.value.map(convEmail => ({
                        subject: convEmail.subject,
                        body: convEmail.body?.content || '',
                        from_email: convEmail.from?.emailAddress?.address || '',
                        date: convEmail.receivedDateTime || new Date().toISOString()
                    }));
                }
            }

            // Fallback: return just this email
            return [{
                subject: email.subject,
                body: email.body?.content || email.bodyPreview || '',
                from_email: email.from?.emailAddress?.address || '',
                date: email.receivedDateTime || new Date().toISOString()
            }];
        } catch (error) {
            console.error('Error getting email chain:', error);
            return [];
        }
    }
};

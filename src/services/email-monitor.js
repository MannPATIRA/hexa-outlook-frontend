/**
 * Email Monitoring Service
 * Automatically monitors for replies to emails in Sent RFQs folder
 * Classifies and organizes them automatically
 */
const EmailMonitor = {
    // Track processed emails to avoid duplicates
    processedEmails: new Set(),
    
    // Monitoring interval (in milliseconds)
    monitoringInterval: 30000, // 30 seconds
    
    // Interval ID for cleanup
    intervalId: null,
    
    // Whether monitoring is active
    isMonitoring: false,

    /**
     * Start monitoring for replies to Sent RFQs
     */
    async startMonitoring() {
        if (this.isMonitoring) {
            console.log('Email monitoring already active');
            return;
        }

        if (!AuthService.isSignedIn()) {
            console.log('Not signed in - email monitoring will start after sign-in');
            return;
        }

        console.log('Starting automatic email monitoring...');
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
        console.log('Email monitoring stopped');
    },

    /**
     * Check for new replies to Sent RFQs
     */
    async checkForReplies() {
        try {
            if (!AuthService.isSignedIn()) {
                return;
            }

            console.log('Checking for replies to Sent RFQs...');

            // Get recent unread emails from Inbox
            const recentEmails = await this.getRecentInboxEmails();
            
            for (const email of recentEmails) {
                // Skip if already processed
                if (this.processedEmails.has(email.id)) {
                    continue;
                }

                // Check if this is a reply to an email with SENT RFQ category
                const isReplyToSentRfq = await this.isReplyToSentRfq(email);
                
                if (isReplyToSentRfq) {
                    console.log(`Found reply to Sent RFQ: ${email.id} - ${email.subject}`);
                    await this.processReply(email, isReplyToSentRfq);
                    this.processedEmails.add(email.id);
                }
            }
        } catch (error) {
            console.error('Error checking for replies:', error);
        }
    },

    /**
     * Get recent emails from Inbox
     */
    async getRecentInboxEmails() {
        try {
            const endpoint = '/me/mailFolders/inbox/messages?' +
                '$select=id,subject,from,receivedDateTime,conversationId,internetMessageId,bodyPreview,isRead' +
                '&$filter=isRead eq false' +
                '&$top=20' +
                '&$orderby=receivedDateTime desc';
            
            const response = await AuthService.graphRequest(endpoint);
            return response.value || [];
        } catch (error) {
            console.error('Error getting recent inbox emails:', error);
            return [];
        }
    },

    /**
     * Check if an email is a reply to an email with SENT RFQ category
     * Returns the parent email info if found, null otherwise
     */
    async isReplyToSentRfq(email) {
        try {
            // Get full email details including conversation ID
            const fullEmail = await AuthService.graphRequest(
                `/me/messages/${email.id}?$select=id,subject,conversationId,internetMessageId,inReplyTo`
            );

            if (!fullEmail.conversationId) {
                return null;
            }

            // Get all emails in the conversation
            const conversationEmails = await AuthService.graphRequest(
                `/me/messages?$filter=conversationId eq '${fullEmail.conversationId}'` +
                `&$select=id,subject,categories,receivedDateTime` +
                `&$orderby=receivedDateTime asc`
            );

            if (!conversationEmails.value || conversationEmails.value.length === 0) {
                return null;
            }

            // Find the original email in the conversation that has SENT RFQ category
            // The original should be the first email with SENT RFQ category
            for (const convEmail of conversationEmails.value) {
                if (convEmail.id === email.id) {
                    continue; // Skip the current email
                }

                // Check if this email has SENT RFQ category
                if (convEmail.categories && 
                    convEmail.categories.some(cat => 
                        cat.toLowerCase().includes('sent rfq')
                    )) {
                    // Found the parent email with SENT RFQ category
                    // Extract material code from the folder path
                    const materialCode = await this.extractMaterialCodeFromEmail(convEmail.id);
                    
                    return {
                        parentEmailId: convEmail.id,
                        parentSubject: convEmail.subject,
                        materialCode: materialCode
                    };
                }
            }

            // Alternative: Check if email is in a Sent RFQs folder
            // Get the email's folder
            const emailWithFolder = await AuthService.graphRequest(
                `/me/messages/${email.id}?$select=id,parentFolderId`
            );

            if (emailWithFolder.parentFolderId) {
                const folderInfo = await this.getFolderInfo(emailWithFolder.parentFolderId);
                if (folderInfo && folderInfo.displayName === 'Sent RFQs') {
                    // Extract material code from parent folder
                    const materialCode = await this.extractMaterialCodeFromParentFolder(emailWithFolder.parentFolderId);
                    return {
                        parentEmailId: null,
                        parentSubject: email.subject,
                        materialCode: materialCode
                    };
                }
            }

            return null;
        } catch (error) {
            console.error('Error checking if reply to Sent RFQ:', error);
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
            return null;
        } catch (error) {
            console.error('Error extracting material code from folder:', error);
            return null;
        }
    },

    /**
     * Process a reply email: classify, categorize, and move to appropriate folder
     */
    async processReply(email, replyInfo) {
        try {
            console.log(`Processing reply email ${email.id}...`);

            // Get full email details with body content
            const fullEmail = await AuthService.graphRequest(
                `/me/messages/${email.id}?$select=id,subject,from,body,receivedDateTime,conversationId`
            );
            
            // Get email chain for classification
            const emailChain = await this.getEmailChain(email.id);
            
            // Extract RFQ ID from subject if possible
            const rfqId = EmailOperations.extractRfqId(fullEmail.subject);

            // Get email body content (prefer HTML, fallback to text)
            let emailBody = '';
            if (fullEmail.body) {
                emailBody = fullEmail.body.content || '';
            }

            // Classify the email
            console.log('Classifying email...');
            const classification = await ApiClient.classifyEmail(
                emailChain,
                {
                    subject: fullEmail.subject || '',
                    body: emailBody,
                    from_email: fullEmail.from?.emailAddress || '',
                    date: fullEmail.receivedDateTime || new Date().toISOString(),
                    in_reply_to: rfqId
                },
                rfqId
            );

            console.log(`Email classified as: ${classification.classification} (confidence: ${classification.confidence})`);

            // Get material code (from reply info or extract from subject/folder)
            let materialCode = replyInfo.materialCode;
            if (!materialCode) {
                // Try to extract from subject (e.g., "RFQ for MAT-12345")
                const match = fullEmail.subject.match(/MAT-\d+/i);
                if (match) {
                    materialCode = match[0];
                }
            }

            if (!materialCode) {
                console.warn('Could not determine material code for email, skipping folder organization');
                return;
            }

            // Determine target folder based on classification
            const targetFolder = FolderManagement.getFolderForClassification(
                materialCode,
                classification.classification,
                classification.sub_classification || null
            );

            console.log(`Moving email to folder: ${targetFolder}`);

            // Ensure folders exist
            await FolderManagement.initializeMaterialFolders(materialCode);

            // Move email to appropriate folder
            await FolderManagement.moveEmailToFolder(email.id, targetFolder);
            console.log(`✓ Moved email to ${targetFolder}`);

            // Apply appropriate category based on classification
            let categoryName = null;
            switch (classification.classification) {
                case 'quote':
                    categoryName = 'QUOTE';
                    break;
                case 'clarification_request':
                    categoryName = 'CLARIFICATION';
                    break;
                case 'engineer_response':
                    categoryName = 'ENGINEER RESPONSE';
                    break;
            }

            if (categoryName) {
                try {
                    await EmailOperations.applyCategoryToEmail(email.id, categoryName);
                    console.log(`✓ Applied category "${categoryName}" to email`);
                } catch (categoryError) {
                    console.error('Failed to apply category (non-critical):', categoryError);
                }
            }

            // Mark email as read
            try {
                await EmailOperations.markAsRead(email.id, true);
            } catch (readError) {
                console.error('Failed to mark email as read (non-critical):', readError);
            }

            console.log(`✓ Successfully processed reply email ${email.id}`);
        } catch (error) {
            console.error('Error processing reply:', error);
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
                        from_email: convEmail.from?.emailAddress || '',
                        date: convEmail.receivedDateTime || new Date().toISOString()
                    }));
                }
            }

            // Fallback: return just this email
            return [{
                subject: email.subject,
                body: email.body?.content || email.bodyPreview || '',
                from_email: email.from?.emailAddress || '',
                date: email.receivedDateTime || new Date().toISOString()
            }];
        } catch (error) {
            console.error('Error getting email chain:', error);
            return [];
        }
    }
};

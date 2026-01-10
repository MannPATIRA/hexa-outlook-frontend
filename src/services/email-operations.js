/**
 * Email Operations Service
 * Handles email reading, composing, and sending
 * Supports both Office.js (drafts) and Graph API (direct send)
 */
const EmailOperations = {
    /**
     * Get current email context from Outlook
     * Only works when an email is selected/open
     */
    async getCurrentEmail() {
        return new Promise((resolve, reject) => {
            if (!Office.context.mailbox.item) {
                reject(new Error('No email selected'));
                return;
            }

            const item = Office.context.mailbox.item;
            const email = {
                id: item.itemId,
                subject: item.subject,
                from: null,
                to: [],
                cc: [],
                receivedDateTime: item.dateTimeCreated,
                conversationId: item.conversationId
            };

            // Get from address
            if (item.from) {
                item.from.getAsync((result) => {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        email.from = {
                            name: result.value.displayName,
                            address: result.value.emailAddress
                        };
                    }
                });
            }

            // Get body
            item.body.getAsync(Office.CoercionType.Html, (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    email.body = result.value;
                }

                // Get plain text too
                item.body.getAsync(Office.CoercionType.Text, (textResult) => {
                    if (textResult.status === Office.AsyncResultStatus.Succeeded) {
                        email.bodyText = textResult.value;
                    }
                    resolve(email);
                });
            });
        });
    },

    /**
     * Get email by ID using Graph API
     */
    async getEmailById(emailId) {
        if (!AuthService.isSignedIn()) {
            throw new Error('Please sign in to read emails');
        }

        return await AuthService.graphRequest(
            `/me/messages/${emailId}?$select=id,subject,from,toRecipients,ccRecipients,body,receivedDateTime,conversationId,internetMessageId`
        );
    },

    /**
     * Create email draft using Office.js (opens compose window)
     * User reviews and sends manually
     * @param {string|string[]|object} to - Recipient(s) email address(es)
     * @param {string} subject - Email subject
     * @param {string} body - Email HTML body
     * @param {Array} attachments - Optional attachments
     */
    async createDraft(to, subject, body, attachments) {
                // Handle if called with options object (backwards compatibility)
                if (to && typeof to === 'object' && !Array.isArray(to) && to.to) {
                    const options = to;
                    to = options.to;
                    subject = options.subject;
                    body = options.body;
                    attachments = options.attachments;
                }

                // Normalize 'to' to array
                let recipients = [];
                if (typeof to === 'string') {
                    recipients = [to];
                } else if (Array.isArray(to)) {
                    recipients = to;
                } else if (to && to.address) {
                    recipients = [to.address];
                }

                if (recipients.length === 0) {
            throw new Error('Recipients are required');
        }

        // Ensure body is a valid HTML string - Office.js REQUIRES htmlBody to be valid HTML
        // Extract body content (handles both string and object formats)
        const bodyContent = this.extractBodyContent(body);
        
        let safeBody = '<div>&nbsp;</div>'; // Default to non-empty HTML
        
        if (bodyContent && typeof bodyContent === 'string' && bodyContent.trim().length > 0) {
            const trimmed = bodyContent.trim();
            // Check if already HTML (starts with < and contains HTML tags)
            if (trimmed.startsWith('<') && (trimmed.includes('</') || trimmed.includes('/>'))) {
                safeBody = trimmed;
                    } else {
                // Convert plain text to HTML - escape special characters and convert newlines
                safeBody = '<div>' + trimmed
                    .replace(/&/g, '&amp;')
                    .replace(/</g, '&lt;')
                    .replace(/>/g, '&gt;')
                    .replace(/"/g, '&quot;')
                    .replace(/'/g, '&#39;')
                    .replace(/\n/g, '<br>') + '</div>';
            }
        }

        // If user is signed in, save draft to Drafts folder using Graph API
        if (AuthService.isSignedIn()) {
            try {
                console.log('createDraft: User signed in, saving draft to Drafts folder via Graph API');
                
                // Save draft to Drafts folder
                const draft = await this.saveDraft({
                    to: recipients,
                    subject: subject || '',
                    body: safeBody,
                    cc: [],
                    attachments: attachments
                });

                console.log('createDraft: Draft saved with ID:', draft.id);
                
                // Now open compose window with the same content (shows in current Outlook window)
                // This allows user to edit and the draft is already saved in Drafts folder
                const mailItem = {
                    toRecipients: recipients.map(email => 
                        typeof email === 'string' ? email : email.address
                    ),
                    subject: (subject && typeof subject === 'string') ? subject : '',
                    htmlBody: safeBody
                };

                if (attachments && Array.isArray(attachments) && attachments.length > 0) {
                    mailItem.attachments = attachments;
                }

                // Open compose window in current Outlook window
                Office.context.mailbox.displayNewMessageForm(mailItem);
                
                return { 
                    status: 'draft_saved_and_opened', 
                    draftId: draft.id,
                    message: 'Draft saved to Drafts folder and opened for editing'
                };
            } catch (error) {
                console.error('createDraft: Error saving draft via Graph API, falling back to Office.js:', error);
                // Fall through to Office.js method
            }
        }

        // Fallback: Use Office.js to open compose window (if not signed in or Graph API failed)
        try {
            console.log('createDraft: Using Office.js compose window');
            
            const mailItem = {
                toRecipients: recipients.map(email => 
                    typeof email === 'string' ? email : email.address
                ),
                subject: (subject && typeof subject === 'string') ? subject : '',
                htmlBody: safeBody
            };

            if (attachments && Array.isArray(attachments) && attachments.length > 0) {
                mailItem.attachments = attachments;
            }

            // Log for debugging
            console.log('createDraft - recipients:', mailItem.toRecipients);
            console.log('createDraft - subject:', mailItem.subject);
            console.log('createDraft - htmlBody length:', safeBody.length);
            console.log('createDraft - htmlBody preview:', safeBody.substring(0, 100) + (safeBody.length > 100 ? '...' : ''));
            
            Office.context.mailbox.displayNewMessageForm(mailItem);
            return { status: 'draft_opened' };
        } catch (error) {
            console.error('createDraft error:', error);
            throw error;
        }
    },

    /**
     * Send email directly using Graph API
     * No user intervention required
     * @param {Object} options - Email options
     * @param {string|Array} options.to - Recipient email(s)
     * @param {string} options.subject - Email subject
     * @param {string} options.body - Email HTML body
     * @param {Array} options.cc - CC recipients (optional)
     * @param {string} options.materialCode - Material code for folder organization (optional)
     * @returns {Promise<Object>} Result with status and sent email ID
     */
    async sendEmail(options) {
        if (!AuthService.isSignedIn()) {
            throw new Error('Please sign in to send emails');
        }

        const message = {
            subject: options.subject,
            body: {
                contentType: 'HTML',
                content: options.body
            },
            toRecipients: options.to.map(email => ({
                emailAddress: {
                    address: typeof email === 'string' ? email : email.address,
                    name: typeof email === 'string' ? undefined : email.name
                }
            }))
        };

        if (options.cc && options.cc.length > 0) {
            message.ccRecipients = options.cc.map(email => ({
                emailAddress: {
                    address: typeof email === 'string' ? email : email.address,
                    name: typeof email === 'string' ? undefined : email.name
                }
            }));
        }

        // Send the email
        console.log('Sending email via Graph API:', {
            to: options.to,
            subject: options.subject,
            hasBody: !!options.body,
            bodyLength: options.body?.length || 0
        });

        try {
            const response = await AuthService.graphRequest('/me/sendMail', {
            method: 'POST',
            body: JSON.stringify({
                message: message,
                saveToSentItems: true
            })
        });
            console.log('Email sent successfully via Graph API. Response:', response);
        } catch (sendError) {
            console.error('Graph API sendMail failed:', sendError);
            console.error('Error details:', {
                message: sendError.message,
                status: sendError.statusCode,
                data: sendError.data
            });
            
            // Provide more helpful error messages
            let errorMessage = 'Failed to send email';
            if (sendError.message) {
                errorMessage = sendError.message;
            } else if (sendError.statusCode === 403) {
                errorMessage = 'Permission denied. Please check Mail.Send permission.';
            } else if (sendError.statusCode === 400) {
                errorMessage = 'Invalid email format or missing required fields.';
            }
            
            throw new Error(errorMessage);
        }

        let sentEmailId = null;

        // If materialCode provided, move email to Sent RFQs folder
        if (options.materialCode) {
            try {
                // Wait for email to appear in Sent Items (can take 3-10 seconds)
                // Try multiple times with increasing delays
                let sentEmail = null;
                const maxRetries = 5;
                const initialDelay = 2000; // Start with 2 seconds
                
                for (let attempt = 0; attempt < maxRetries && !sentEmail; attempt++) {
                    const delay = initialDelay + (attempt * 1000); // Increase delay each attempt
                    console.log(`Waiting ${delay}ms before searching for sent email (attempt ${attempt + 1}/${maxRetries})...`);
                    await new Promise(resolve => setTimeout(resolve, delay));

                    // Try to find the sent email - use multiple strategies
                    // Strategy 1: Find by exact subject match
                    const sentEmailsBySubject = await this.getSentEmails({
                        subject: options.subject,
                        top: 10
                    });

                    if (sentEmailsBySubject.length > 0) {
                        // Filter by recipient to ensure we get the right one
                        const recipientEmail = typeof options.to[0] === 'string' 
                            ? options.to[0] 
                            : options.to[0].address;
                        
                        for (const email of sentEmailsBySubject) {
                            const toRecipients = email.toRecipients || [];
                            const matchesRecipient = toRecipients.some(recipient => 
                                recipient.emailAddress?.address?.toLowerCase() === recipientEmail.toLowerCase()
                            );
                            
                            if (matchesRecipient) {
                                sentEmail = email;
                                console.log(`Found sent email on attempt ${attempt + 1} by subject + recipient match`);
                                break;
                            }
                        }
                        
                        // If no recipient match, use the most recent one
                        if (!sentEmail && sentEmailsBySubject.length > 0) {
                            sentEmail = sentEmailsBySubject[0];
                            console.log(`Found sent email on attempt ${attempt + 1} by subject match (no recipient match)`);
                        }
                    }

                    // Strategy 2: If not found, try getting most recent emails and match by recipient
                    if (!sentEmail) {
                        const recentEmails = await this.getSentEmails({ top: 20 });
                        const recipientEmail = typeof options.to[0] === 'string' 
                            ? options.to[0] 
                            : options.to[0].address;
                        
                        for (const email of recentEmails) {
                            const toRecipients = email.toRecipients || [];
                            const matchesRecipient = toRecipients.some(recipient => 
                                recipient.emailAddress?.address?.toLowerCase() === recipientEmail.toLowerCase()
                            );
                            const matchesSubject = email.subject === options.subject;
                            
                            if (matchesRecipient && matchesSubject) {
                                sentEmail = email;
                                console.log(`Found sent email on attempt ${attempt + 1} by recent emails + recipient + subject match`);
                                break;
                            }
                        }
                    }

                    // Strategy 3: If still not found and this is the last attempt, use the most recent email with matching subject
                    if (!sentEmail && attempt === maxRetries - 1) {
                        console.log('Last attempt - trying to find by subject only from recent emails');
                        const recentEmails = await this.getSentEmails({ top: 10 });
                        for (const email of recentEmails) {
                            if (email.subject === options.subject) {
                                sentEmail = email;
                                console.log(`Found sent email on final attempt by subject match only`);
                                break;
                            }
                        }
                    }
                }

                if (sentEmail) {
                    sentEmailId = sentEmail.id;
                    console.log(`Found sent email with ID: ${sentEmailId}`);
                    console.log(`Email details:`, {
                        id: sentEmail.id,
                        subject: sentEmail.subject,
                        toRecipients: sentEmail.toRecipients || []
                    });

                    // Ensure folder structure exists
                    try {
                        console.log(`Initializing folder structure for: ${options.materialCode}`);
                        await FolderManagement.initializeMaterialFolders(options.materialCode);
                        console.log('Folder structure initialized successfully');
                    } catch (folderError) {
                        console.error('Failed to initialize folders:', folderError);
                        throw folderError; // Re-throw so we know folder creation failed
                    }

                    // Move to Sent RFQs folder FIRST
                    const folderPath = `${options.materialCode}/${Config.FOLDERS.SENT_RFQS}`;
                    console.log(`Attempting to move email to: ${folderPath}`);
                    let movedEmailId = sentEmailId;
                    try {
                        const moveResult = await FolderManagement.moveEmailToFolder(sentEmailId, folderPath);
                        console.log(`✓ Successfully moved sent email ${sentEmailId} to ${folderPath}`);
                        
                        // Use the email ID from move result if available, otherwise use original
                        movedEmailId = moveResult?.id || sentEmailId;
                        console.log(`Moved email ID: ${movedEmailId}`);
                        
                        // Wait a moment for the move to complete
                        await new Promise(resolve => setTimeout(resolve, 1000));
                        
                        // Apply "SENT RFQ" category tag to the email AFTER moving
                        // This ensures the category is applied to the email in its final location
                        try {
                            console.log(`Applying SENT RFQ category to moved email ${movedEmailId}...`);
                            await this.applyCategoryToEmail(movedEmailId, 'SENT RFQ');
                            console.log(`✓ Successfully applied SENT RFQ category to email ${movedEmailId}`);
                            
                            // Verify the category was applied by checking again
                            await new Promise(resolve => setTimeout(resolve, 500));
                            const verifyEmail = await AuthService.graphRequest(`/me/messages/${movedEmailId}?$select=id,categories`);
                            if (verifyEmail.categories && verifyEmail.categories.some(cat => 
                                cat.toLowerCase().includes('sent rfq')
                            )) {
                                console.log(`✓ Category verified on email ${movedEmailId}:`, verifyEmail.categories);
                            } else {
                                console.warn(`⚠ Category not found after verification. Retrying...`);
                                // Retry once more
                                await this.applyCategoryToEmail(movedEmailId, 'SENT RFQ');
                            }
                        } catch (categoryError) {
                            console.error('Failed to apply category (non-critical):', categoryError);
                            console.error('Category error details:', {
                                emailId: movedEmailId,
                                error: categoryError.message
                            });
                            // Don't fail the operation if category fails, but log it
                        }
                    } catch (moveError) {
                        console.error('✗ Failed to move email:', moveError);
                        console.error('Move error details:', {
                            emailId: sentEmailId,
                            folderPath: folderPath,
                            error: moveError.message
                        });
                        throw moveError; // Re-throw so we can see the error
                    }
                } else {
                    console.warn('⚠ Could not find sent email to move to folder');
                    console.warn('Search details:', {
                        subject: options.subject,
                        to: options.to,
                        attempts: maxRetries
                    });
                    console.warn('This may be due to timing - email may appear in folder after a few seconds');
                }
            } catch (error) {
                console.error('❌ FAILED to move sent email to folder:', error);
                console.error('Error details:', {
                    message: error.message,
                    stack: error.stack,
                    materialCode: options.materialCode,
                    subject: options.subject,
                    to: options.to
                });
                // Don't fail the send operation if moving fails, but log it prominently
                // The email was sent successfully, just not moved to the folder
            }
        }

        return { 
            status: 'sent',
            sentEmailId: sentEmailId
        };
    },

    /**
     * Get sent emails from Sent Items folder
     * @param {Object} options - Search options
     * @param {string} options.subject - Filter by subject (exact match)
     * @param {number} options.top - Number of results to return (default: 10)
     * @returns {Promise<Array>} Array of sent email objects
     */
    async getSentEmails(options = {}) {
        if (!AuthService.isSignedIn()) {
            throw new Error('Please sign in to get sent emails');
        }

        let endpoint = '/me/mailFolders/sentitems/messages';
        const params = [];

        // Note: OData filter on subject can be unreliable, so we'll get more results and filter in JS
        // Don't use filter for now - get more emails and filter in JavaScript
        // if (options.subject) {
        //     const escapedSubject = options.subject.replace(/'/g, "''");
        //     params.push(`$filter=subject eq '${escapedSubject}'`);
        // }

        params.push('$orderby=sentDateTime desc,createdDateTime desc');
        // Get more results if filtering by subject (we'll filter in JS)
        params.push(`$top=${options.subject ? (options.top || 50) : (options.top || 10)}`);
        params.push('$select=id,subject,toRecipients,createdDateTime,sentDateTime,internetMessageId');

        if (params.length > 0) {
            endpoint += '?' + params.join('&');
        }

        try {
            console.log(`Querying sent emails: ${endpoint}`);
            const response = await AuthService.graphRequest(endpoint);
            let emails = response.value || [];
            console.log(`Found ${emails.length} sent emails from API`);
            
            // Filter by subject in JavaScript if provided (more reliable than OData filter)
            if (options.subject && emails.length > 0) {
                const filtered = emails.filter(email => 
                    email.subject && email.subject.trim() === options.subject.trim()
                );
                console.log(`Filtered to ${filtered.length} emails matching subject: "${options.subject}"`);
                emails = filtered;
            }
            
            if (emails.length > 0) {
                console.log('Sample email:', {
                    id: emails[0].id,
                    subject: emails[0].subject,
                    sentDateTime: emails[0].sentDateTime,
                    toRecipients: emails[0].toRecipients?.map(r => r.emailAddress?.address) || []
                });
            }
            return emails;
        } catch (error) {
            console.error('Error getting sent emails:', error);
            // If filter fails, try without filter and filter in JavaScript
            if (options.subject && (error.message?.includes('filter') || error.message?.includes('Invalid'))) {
                console.log('Filter query failed, retrying without filter and filtering in code...');
                try {
                    const allEmails = await this.getSentEmails({ top: options.top || 50 });
                    // Filter by subject in JavaScript
                    const filtered = allEmails.filter(email => 
                        email.subject && email.subject.includes(options.subject)
                    );
                    console.log(`Filtered to ${filtered.length} emails matching subject`);
                    return filtered;
                } catch (retryError) {
                    console.error('Retry also failed:', retryError);
                    throw error; // Throw original error
                }
            }
            throw error;
        }
    },

    /**
     * Get or create a category by name
     * @param {string} categoryName - Name of the category
     * @param {string} color - Color for the category (Preset0-Preset24, or null for default)
     * @returns {Promise<string>} Category display name (used to apply to emails)
     */
    async getOrCreateCategory(categoryName, color = 'Preset0') {
        if (!AuthService.isSignedIn()) {
            throw new Error('Please sign in to manage categories');
        }

        try {
            // Get all existing categories
            const response = await AuthService.graphRequest('/me/outlook/masterCategories');
            const categories = response.value || [];
            console.log(`Found ${categories.length} existing categories`);
            
            // Check if category already exists (case-insensitive)
            const existing = categories.find(cat => 
                cat.displayName && cat.displayName.toLowerCase() === categoryName.toLowerCase()
            );
            
            if (existing) {
                console.log(`Category "${categoryName}" already exists with displayName: "${existing.displayName}"`);
                return existing.displayName; // Return the exact displayName from the API
            }

            // Create new category
            console.log(`Creating category "${categoryName}" with color ${color}`);
            const newCategory = await AuthService.graphRequest('/me/outlook/masterCategories', {
                method: 'POST',
                body: JSON.stringify({
                    displayName: categoryName,
                    color: color
                })
            });
            
            console.log(`✓ Created category "${categoryName}" with ID: ${newCategory.id}, displayName: "${newCategory.displayName}"`);
            return newCategory.displayName;
        } catch (error) {
            console.error('Error getting/creating category:', error);
            // If creation fails, try to use the name anyway (might work if category exists)
            return categoryName;
        }
    },

    /**
     * Apply a category to an email
     * @param {string} emailId - The ID of the email
     * @param {string} categoryName - Name of the category to apply
     */
    async applyCategoryToEmail(emailId, categoryName) {
        if (!AuthService.isSignedIn()) {
            throw new Error('Please sign in to apply categories');
        }

        try {
            // First, ensure the category exists
            // Preset0 = Red, Preset1 = Orange, Preset2 = Brown, Preset3 = Yellow, Preset4 = Green, Preset5 = Teal, Preset6 = Blue, Preset7 = Purple, Preset8 = Pink, Preset9 = Gray
            // Using Preset6 (Blue) for SENT RFQ category - light blue color
            const categoryDisplayName = await this.getOrCreateCategory(categoryName, 'Preset6'); // Light blue color
            console.log(`Category display name: "${categoryDisplayName}"`);
            
            // Get current email to see existing categories
            const email = await AuthService.graphRequest(`/me/messages/${emailId}?$select=id,categories`);
            console.log(`Current email categories:`, email.categories);
            const currentCategories = email.categories || [];
            
            // Add category if not already present
            if (!currentCategories.includes(categoryDisplayName)) {
                const updatedCategories = [...currentCategories, categoryDisplayName];
                console.log(`Updating email ${emailId} with categories:`, updatedCategories);
                
                // Update email with new categories
                const updateResponse = await AuthService.graphRequest(`/me/messages/${emailId}`, {
                    method: 'PATCH',
                    body: JSON.stringify({
                        categories: updatedCategories
                    })
                });
                
                // Verify the category was applied by fetching the email again
                const verifyEmail = await AuthService.graphRequest(`/me/messages/${emailId}?$select=id,categories`);
                console.log(`Verified email categories after update:`, verifyEmail.categories);
                
                if (verifyEmail.categories && verifyEmail.categories.includes(categoryDisplayName)) {
                    console.log(`✓ Successfully applied category "${categoryDisplayName}" to email ${emailId}`);
                } else {
                    console.warn(`⚠ Category "${categoryDisplayName}" may not have been applied correctly`);
                }
            } else {
                console.log(`Category "${categoryDisplayName}" already applied to email ${emailId}`);
            }
        } catch (error) {
            console.error('Error applying category to email:', error);
            console.error('Error details:', {
                emailId: emailId,
                categoryName: categoryName,
                errorMessage: error.message,
                errorStack: error.stack
            });
            throw error;
        }
    },

    /**
     * Delete a draft by ID
     * @param {string} draftId - The ID of the draft to delete
     */
    async deleteDraft(draftId) {
        if (!AuthService.isSignedIn()) {
            throw new Error('Please sign in to delete drafts');
        }

        if (!draftId) {
            throw new Error('Draft ID is required');
        }

        try {
            console.log(`Deleting draft: ${draftId}`);
            await AuthService.graphRequest(`/me/messages/${draftId}`, {
                method: 'DELETE'
            });
            console.log(`✓ Successfully deleted draft: ${draftId}`);
            return { status: 'deleted' };
        } catch (error) {
            console.error('Failed to delete draft:', error);
            // Don't throw - draft deletion failure shouldn't fail the send operation
            return { status: 'delete_failed', error: error.message };
        }
    },

    /**
     * Create and save draft using Graph API (doesn't open compose window)
     */
    async saveDraft(options) {
        if (!AuthService.isSignedIn()) {
            throw new Error('Please sign in to save drafts');
        }

        const message = {
            subject: options.subject,
            body: {
                contentType: 'HTML',
                content: options.body
            },
            toRecipients: options.to.map(email => ({
                emailAddress: {
                    address: typeof email === 'string' ? email : email.address
                }
            }))
        };

        if (options.cc && options.cc.length > 0) {
            message.ccRecipients = options.cc.map(email => ({
                emailAddress: {
                    address: typeof email === 'string' ? email : email.address
                }
            }));
        }

        const draft = await AuthService.graphRequest('/me/messages', {
            method: 'POST',
            body: JSON.stringify(message)
        });

        return draft;
    },

    /**
     * Open an existing draft by ID
     * Gets the draft from Graph API and opens it in Outlook compose window
     * @param {string} draftId - The ID of the draft to open
     */
    async openDraft(draftId) {
        if (!AuthService.isSignedIn()) {
            throw new Error('Please sign in to open drafts');
        }

        try {
            // Get the draft by ID
            const draft = await AuthService.graphRequest(`/me/messages/${draftId}?$select=id,subject,toRecipients,body,ccRecipients`);

            // Extract email details
            const toRecipients = (draft.toRecipients || []).map(recipient => 
                recipient.emailAddress.address
            );
            const subject = draft.subject || '';
            const htmlBody = draft.body?.content || '<div>&nbsp;</div>';

            // Open compose window with draft content
            const mailItem = {
                toRecipients: toRecipients,
                subject: subject,
                htmlBody: htmlBody
            };

            if (draft.ccRecipients && draft.ccRecipients.length > 0) {
                mailItem.ccRecipients = draft.ccRecipients.map(recipient => 
                    recipient.emailAddress.address
                );
            }

            Office.context.mailbox.displayNewMessageForm(mailItem);
            return { status: 'draft_opened', draftId: draftId };
        } catch (error) {
            console.error('Error opening draft:', error);
            throw error;
        }
    },

    /**
     * Reply to an email using Graph API
     */
    async replyToEmail(emailId, replyContent, replyAll = false) {
        if (!AuthService.isSignedIn()) {
            throw new Error('Please sign in to reply to emails');
        }

        const endpoint = replyAll 
            ? `/me/messages/${emailId}/replyAll`
            : `/me/messages/${emailId}/reply`;

        await AuthService.graphRequest(endpoint, {
            method: 'POST',
            body: JSON.stringify({
                message: {
                    body: {
                        contentType: 'HTML',
                        content: replyContent
                    }
                }
            })
        });

        return { status: 'replied' };
    },

    /**
     * Forward an email using Graph API
     */
    async forwardEmail(emailId, toRecipients, comment = '') {
        if (!AuthService.isSignedIn()) {
            throw new Error('Please sign in to forward emails');
        }

        await AuthService.graphRequest(`/me/messages/${emailId}/forward`, {
            method: 'POST',
            body: JSON.stringify({
                comment: comment,
                toRecipients: toRecipients.map(email => ({
                    emailAddress: {
                        address: typeof email === 'string' ? email : email.address
                    }
                }))
            })
        });

        return { status: 'forwarded' };
    },

    /**
     * Create RFQ email - supports both draft and direct send
     */
    async createRfqEmail(supplierEmail, supplierName, rfqContent, subject, sendDirectly = false) {
        const options = {
            to: [{ address: supplierEmail, name: supplierName }],
            subject: subject,
            body: rfqContent
        };

        if (sendDirectly && AuthService.isSignedIn()) {
            return await this.sendEmail(options);
        } else {
            return this.createDraft(options);
        }
    },

    /**
     * Forward to engineering team
     */
    async forwardToEngineering(emailId, engineeringEmail, notes = '') {
        const comment = notes 
            ? `<p><strong>Procurement Notes:</strong></p><p>${notes}</p><hr/>`
            : '<p>Please review the technical query below and provide your response.</p><hr/>';

        if (AuthService.isSignedIn()) {
            return await this.forwardEmail(emailId, [engineeringEmail], comment);
        } else {
            // Fallback to Office.js draft
            const email = await this.getCurrentEmail();
            return this.createDraft({
                to: [engineeringEmail],
                subject: `FW: ${email.subject}`,
                body: `${comment}<br/><br/>${email.body}`
            });
        }
    },

    /**
     * Send clarification response
     */
    async sendClarificationResponse(originalEmailId, responseContent, sendDirectly = false) {
        if (sendDirectly && AuthService.isSignedIn()) {
            return await this.replyToEmail(originalEmailId, responseContent);
        } else {
            // Open reply draft via Office.js
            const email = await this.getCurrentEmail();
            return this.createDraft({
                to: [email.from],
                subject: `RE: ${email.subject}`,
                body: responseContent
            });
        }
    },

    /**
     * Get emails from inbox with optional filters
     */
    async getInboxEmails(options = {}) {
        if (!AuthService.isSignedIn()) {
            throw new Error('Please sign in to get emails');
        }

        let endpoint = '/me/mailFolders/inbox/messages';
        const params = [];

        if (options.top) {
            params.push(`$top=${options.top}`);
        } else {
            params.push('$top=50');
        }

        if (options.filter) {
            params.push(`$filter=${encodeURIComponent(options.filter)}`);
        }

        if (options.search) {
            params.push(`$search="${encodeURIComponent(options.search)}"`);
        }

        params.push('$select=id,subject,from,receivedDateTime,bodyPreview,isRead');
        params.push('$orderby=receivedDateTime desc');

        endpoint += '?' + params.join('&');

        const response = await AuthService.graphRequest(endpoint);
        return response.value || [];
    },

    /**
     * Mark email as read
     */
    async markAsRead(emailId, isRead = true) {
        if (!AuthService.isSignedIn()) {
            throw new Error('Please sign in to update emails');
        }

        await AuthService.graphRequest(`/me/messages/${emailId}`, {
            method: 'PATCH',
            body: JSON.stringify({ isRead: isRead })
        });

        return { status: 'updated' };
    },

    /**
     * Search for emails by subject or sender
     */
    async searchEmails(query, folder = 'inbox') {
        if (!AuthService.isSignedIn()) {
            throw new Error('Please sign in to search emails');
        }

        const endpoint = `/me/mailFolders/${folder}/messages?$search="${encodeURIComponent(query)}"&$top=25&$select=id,subject,from,receivedDateTime,bodyPreview`;

        const response = await AuthService.graphRequest(endpoint);
        return response.value || [];
    },

    /**
     * Get conversation thread
     */
    async getConversation(conversationId) {
        if (!AuthService.isSignedIn()) {
            throw new Error('Please sign in to get conversations');
        }

        const endpoint = `/me/messages?$filter=conversationId eq '${conversationId}'&$orderby=receivedDateTime asc&$select=id,subject,from,receivedDateTime,body`;

        const response = await AuthService.graphRequest(endpoint);
        return response.value || [];
    },

    /**
     * Extract body content from various formats (string, object, etc.)
     * Handles cases where backend returns body as an object
     */
    extractBodyContent(body) {
        if (!body) return '';
        
        // If it's already a string, return it
        if (typeof body === 'string') {
            return body;
        }
        
        // If it's an object, try to extract meaningful content
        if (typeof body === 'object') {
            console.log('extractBodyContent: body is object, type:', body.constructor?.name, 'keys:', Object.keys(body));
            // Try common property names
            if (body.content) {
                console.log('extractBodyContent: found body.content');
                return String(body.content);
            }
            if (body.text) {
                console.log('extractBodyContent: found body.text');
                return String(body.text);
            }
            if (body.html) {
                console.log('extractBodyContent: found body.html');
                return String(body.html);
            }
            if (body.body) {
                console.log('extractBodyContent: found body.body');
                return String(body.body);
            }
            
            // Try to build from structured object (like from backend)
            if (body.greeting || body.introduction || body.material_details || body.closing) {
                console.log('extractBodyContent: building from structured object');
                let text = '';
                if (body.greeting) text += body.greeting + '\n\n';
                if (body.introduction) text += body.introduction + '\n\n';
                if (body.material_details) {
                    if (typeof body.material_details === 'object') {
                        text += 'Material Details:\n';
                        for (const [key, value] of Object.entries(body.material_details)) {
                            text += `  - ${key}: ${value}\n`;
                        }
                    } else {
                        text += 'Material Details: ' + String(body.material_details) + '\n';
                    }
                }
                if (body.closing) text += '\n' + body.closing;
                const result = text.trim();
                console.log('extractBodyContent: built text length:', result.length);
                return result;
            }
            
            // Last resort: try JSON.stringify for objects with meaningful content
            try {
                const jsonStr = JSON.stringify(body);
                // If JSON is reasonable length and not just empty object/array
                if (jsonStr.length > 2 && jsonStr !== '{}' && jsonStr !== '[]') {
                    // Try to extract text from JSON
                    return jsonStr;
                }
            } catch (e) {
                // Ignore JSON errors
            }
        }
        
        // Fallback: convert to string
        return String(body);
    },

    /**
     * Format RFQ body as plain text for editing
     */
    formatRfqBodyAsText(htmlBody) {
        if (!htmlBody) return '';
        
        // Extract content first (handles objects)
        const bodyContent = this.extractBodyContent(htmlBody);
        
        // If it's already plain text, return as-is
        if (typeof bodyContent !== 'string') return '';
        
        // Simple HTML to text conversion
        let text = bodyContent
            // Replace <br> and <p> with newlines
            .replace(/<br\s*\/?>/gi, '\n')
            .replace(/<\/p>/gi, '\n\n')
            .replace(/<p[^>]*>/gi, '')
            // Replace common block elements with newlines
            .replace(/<\/div>/gi, '\n')
            .replace(/<div[^>]*>/gi, '')
            .replace(/<\/li>/gi, '\n')
            .replace(/<li[^>]*>/gi, '• ')
            // Remove other HTML tags
            .replace(/<[^>]+>/g, '')
            // Decode common HTML entities
            .replace(/&nbsp;/gi, ' ')
            .replace(/&amp;/gi, '&')
            .replace(/&lt;/gi, '<')
            .replace(/&gt;/gi, '>')
            .replace(/&quot;/gi, '"')
            // Clean up extra whitespace
            .replace(/\n\s*\n\s*\n/g, '\n\n')
            .trim();
        
        return text;
    },

    /**
     * Format plain text as HTML for email body
     * Always returns valid HTML (never empty)
     */
    formatTextAsHtml(text) {
        // Convert to string if not already
        let textString = '';
        if (text !== null && text !== undefined) {
            textString = typeof text === 'string' ? text : String(text);
        }
        
        // If empty, return non-empty HTML div
        if (!textString || textString.trim().length === 0) {
            return '<div>&nbsp;</div>';
        }
        
        // Escape HTML entities
        let html = textString
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#39;')
            // Convert newlines to <br>
            .replace(/\n/g, '<br>');
        
        return '<div>' + html + '</div>';
    },

    /**
     * Extract RFQ ID from email subject
     * Looks for patterns like "RFQ for MAT-12345" or "RFQ-12345"
     */
    extractRfqId(subject) {
        if (!subject) return null;
        
        // Try to extract from subject patterns
        // Pattern 1: "RFQ for MAT-12345" -> extract "MAT-12345"
        const matMatch = subject.match(/MAT-\d+/i);
        if (matMatch) {
            return matMatch[0];
        }
        
        // Pattern 2: "RFQ-12345" or "RFQ 12345"
        const rfqMatch = subject.match(/RFQ[- ]?(\d+)/i);
        if (rfqMatch) {
            return rfqMatch[1];
        }
        
        return null;
    },

    /**
     * Get email chain for classification
     */
    async getEmailChain() {
        try {
            if (!Office.context.mailbox.item) {
                return [];
            }

            const item = Office.context.mailbox.item;
            const conversationId = item.conversationId;

            if (!conversationId || !AuthService.isSignedIn()) {
                // Fallback: return just current email
                const email = await this.getCurrentEmail();
                return [{
                    subject: email.subject || '',
                    body: email.bodyText || email.body || '',
                    from_email: email.from?.address || email.from || '',
                    date: email.receivedDateTime || new Date().toISOString()
                }];
            }

            // Get conversation emails via Graph API
            const conversationEmails = await AuthService.graphRequest(
                `/me/messages?$filter=conversationId eq '${conversationId}'` +
                `&$select=id,subject,from,body,receivedDateTime` +
                `&$orderby=receivedDateTime asc`
            );

            if (conversationEmails.value && conversationEmails.value.length > 0) {
                return conversationEmails.value.map(convEmail => ({
                    subject: convEmail.subject || '',
                    body: convEmail.body?.content || '',
                    from_email: convEmail.from?.emailAddress || '',
                    date: convEmail.receivedDateTime || new Date().toISOString()
                }));
            }

            // Fallback: return just current email
            const email = await this.getCurrentEmail();
            return [{
                subject: email.subject || '',
                body: email.bodyText || email.body || '',
                from_email: email.from?.address || email.from || '',
                date: email.receivedDateTime || new Date().toISOString()
            }];
        } catch (error) {
            console.error('Error getting email chain:', error);
            return [];
        }
    }
};
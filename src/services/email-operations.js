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
        let internetMessageId = null;

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
                    internetMessageId = sentEmail.internetMessageId || null;
                    console.log(`Found sent email with ID: ${sentEmailId}`);
                    console.log(`Email details:`, {
                        id: sentEmail.id,
                        subject: sentEmail.subject,
                        toRecipients: sentEmail.toRecipients || [],
                        internetMessageId: internetMessageId
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

                    // Move to Sent RFQs folder
                    // The folder category tag is automatically applied by moveEmailToFolder
                    const folderPath = `${options.materialCode}/${Config.FOLDERS.SENT_RFQS}`;
                    console.log(`Attempting to move email to: ${folderPath}`);
                    let movedEmailId = sentEmailId;
                    try {
                        const moveResult = await FolderManagement.moveEmailToFolder(sentEmailId, folderPath);
                        console.log(`✓ Successfully moved sent email ${sentEmailId} to ${folderPath}`);
                        
                        // Use the email ID from move result if available, otherwise use original
                        movedEmailId = moveResult?.id || sentEmailId;
                        console.log(`Moved email ID: ${movedEmailId} (folder category auto-applied)`);
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
            sentEmailId: sentEmailId,
            internetMessageId: internetMessageId
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
     * @param {string} color - Optional color preset (default: Preset6/Blue)
     *   Preset0 = Red, Preset1 = Orange, Preset2 = Brown, Preset3 = Yellow,
     *   Preset4 = Green, Preset5 = Teal, Preset6 = Blue, Preset7 = Purple,
     *   Preset8 = Pink, Preset9 = Gray
     */
    async applyCategoryToEmail(emailId, categoryName, color = 'Preset6') {
        if (!AuthService.isSignedIn()) {
            throw new Error('Please sign in to apply categories');
        }

        try {
            // First, ensure the category exists with the specified color
            const categoryDisplayName = await this.getOrCreateCategory(categoryName, color);
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
     * Delete an email by ID
     * @param {string} emailId - The ID of the email to delete
     */
    async deleteEmail(emailId) {
        if (!AuthService.isSignedIn()) {
            throw new Error('Please sign in to delete emails');
        }

        if (!emailId) {
            throw new Error('Email ID is required');
        }

        try {
            console.log(`Deleting email: ${emailId}`);
            await AuthService.graphRequest(`/me/messages/${emailId}`, {
                method: 'DELETE'
            });
            console.log(`✓ Successfully deleted email: ${emailId}`);
            return { status: 'deleted' };
        } catch (error) {
            console.error('Failed to delete email:', error);
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
     * Check if an email is from Microsoft Outlook
     * @param {Object} email - Email object with from field
     * @returns {boolean} True if email is from Microsoft Outlook
     */
    isFromMicrosoftOutlook(email) {
        if (!email || !email.from) {
            return false;
        }

        const fromAddress = (email.from?.emailAddress?.address || '').toLowerCase();
        const fromName = (email.from?.emailAddress?.name || '').toLowerCase();

        // Check email domain
        const microsoftDomains = [
            '@microsoft.com',
            '@outlook.com',
            '@office.com',
            '@office365.com',
            '@microsoftonline.com'
        ];

        const isMicrosoftDomain = microsoftDomains.some(domain => 
            fromAddress.includes(domain)
        );

        // Check sender name for Microsoft Outlook indicators
        const microsoftNamePatterns = [
            'microsoft outlook',
            'microsoft office',
            'outlook',
            'office 365',
            'microsoft',
            'exchange online'
        ];

        const isMicrosoftName = microsoftNamePatterns.some(pattern => 
            fromName.includes(pattern)
        );

        return isMicrosoftDomain || isMicrosoftName;
    },

    /**
     * Create and save draft using Graph API (doesn't open compose window)
     * @param {Object} options - Draft options
     * @param {string|Array} options.to - Recipient email(s)
     * @param {string} options.subject - Email subject
     * @param {string} options.body - Email HTML body
     * @param {Array} options.cc - CC recipients (optional)
     * @param {Array} options.attachments - Attachments in Graph API format (optional)
     * @returns {Promise<Object>} Created draft object
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

        // Create the draft first
        const draft = await AuthService.graphRequest('/me/messages', {
            method: 'POST',
            body: JSON.stringify(message)
        });

        // #region agent log
        fetch('http://127.0.0.1:7248/ingest/c8aaba02-7147-41b9-988d-15ca39db2160',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({location:'email-operations.js:787',message:'saveDraft called with attachments check',data:{hasAttachments:!!options.attachments,isArray:Array.isArray(options.attachments),attachmentsLength:options.attachments?.length,draftId:draft.id},timestamp:Date.now(),sessionId:'debug-session',runId:'run3',hypothesisId:'F'})}).catch(()=>{});
        // #endregion
        
        // Validate attachments before proceeding
        if (!options.attachments || !Array.isArray(options.attachments) || options.attachments.length === 0) {
            console.warn(`[saveDraft] No attachments provided for draft ${draft.id} - draft will be created without attachments`);
            console.warn(`[saveDraft] Attachments value:`, options.attachments);
        }
        
        // Add attachments if provided
        if (options.attachments && Array.isArray(options.attachments) && options.attachments.length > 0) {
            // Helper to detect STEP files
            const isStepFile = (filename) => {
                const ext = (filename || '').split('.').pop().toLowerCase();
                return ext === 'step' || ext === 'stp';
            };
            
            // Count STEP files before upload
            const stepFilesToUpload = options.attachments.filter(a => isStepFile(a.name));
            const pdfFilesToUpload = options.attachments.filter(a => (a.name || '').toLowerCase().endsWith('.pdf'));
            
            console.log(`Adding ${options.attachments.length} attachment(s) to draft ${draft.id}`);
            console.log(`  📎 Upload summary: ${pdfFilesToUpload.length} PDF(s), ${stepFilesToUpload.length} STEP file(s)`);
            
            if (stepFilesToUpload.length > 0) {
                console.log(`  🔧 STEP files to upload:`, stepFilesToUpload.map(a => {
                    const size = a.contentBytes ? (a.contentBytes.length / 1024).toFixed(2) : 0;
                    return `${a.name} (${size} KB base64)`;
                }));
            }
            
            console.log('Attachment details:', options.attachments.map(a => ({ name: a.name, hasContent: !!a.contentBytes })));
            
            // #region agent log
            fetch('http://127.0.0.1:7248/ingest/c8aaba02-7147-41b9-988d-15ca39db2160',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({location:'email-operations.js:788',message:'Starting attachment upload to draft',data:{draftId:draft.id,attachmentsCount:options.attachments.length,stepFilesCount:stepFilesToUpload.length,attachments:options.attachments.map(a=>({name:a.name,hasContentBytes:!!a.contentBytes,contentBytesLength:a.contentBytes?.length,contentType:a.contentType,isStepFile:isStepFile(a.name)}))},timestamp:Date.now(),sessionId:'debug-session',runId:'run3',hypothesisId:'F'})}).catch(()=>{});
            // #endregion
            
            let successfulUploads = 0;
            let failedUploads = 0;
            let stepFilesUploaded = 0;
            let stepFilesFailed = 0;
            
            for (let i = 0; i < options.attachments.length; i++) {
                const attachment = options.attachments[i];
                const isStep = isStepFile(attachment.name);
                
                try {
                    // Validate attachment has required fields
                    if (!attachment.contentBytes || attachment.contentBytes.length === 0) {
                        if (isStep) {
                            console.error(`🔧 [STEP FILE] CRITICAL: ${attachment.name} has no contentBytes!`);
                        }
                        // #region agent log
                        fetch('http://127.0.0.1:7248/ingest/c8aaba02-7147-41b9-988d-15ca39db2160',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({location:'email-operations.js:796',message:'Attachment validation failed - no contentBytes',data:{attachmentName:attachment.name,draftId:draft.id,isStepFile:isStep},timestamp:Date.now(),sessionId:'debug-session',runId:'run3',hypothesisId:'F'})}).catch(()=>{});
                        // #endregion
                        throw new Error(`Attachment ${attachment.name} has no contentBytes`);
                    }
                    
                    if (!attachment.name) {
                        // #region agent log
                        fetch('http://127.0.0.1:7248/ingest/c8aaba02-7147-41b9-988d-15ca39db2160',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({location:'email-operations.js:801',message:'Attachment validation failed - no name',data:{draftId:draft.id,isStepFile:isStep},timestamp:Date.now(),sessionId:'debug-session',runId:'run3',hypothesisId:'F'})}).catch(()=>{});
                        // #endregion
                        throw new Error(`Attachment missing name`);
                    }
                    
                    // Ensure attachment has the correct format for Graph API
                    const attachmentPayload = {
                        '@odata.type': '#microsoft.graph.fileAttachment',
                        name: attachment.name,
                        contentType: attachment.contentType || (isStep ? 'application/octet-stream' : 'application/pdf'),
                        contentBytes: attachment.contentBytes
                    };
                    
                    if (isStep) {
                        const sizeKB = (attachmentPayload.contentBytes.length / 1024).toFixed(2);
                        console.log(`🔧 [STEP FILE] Uploading ${i + 1}/${options.attachments.length}: ${attachment.name} (${sizeKB} KB base64)`);
                    } else {
                        console.log(`[saveDraft] Uploading attachment ${i + 1}/${options.attachments.length}: ${attachment.name}`);
                        console.log(`[saveDraft] Content size: ${attachmentPayload.contentBytes.length} base64 chars`);
                    }
                    
                    // #region agent log
                    fetch('http://127.0.0.1:7248/ingest/c8aaba02-7147-41b9-988d-15ca39db2160',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({location:'email-operations.js:812',message:'Uploading attachment to Graph API',data:{draftId:draft.id,attachmentName:attachment.name,contentBytesLength:attachmentPayload.contentBytes.length,contentType:attachmentPayload.contentType,attachmentIndex:i+1,totalAttachments:options.attachments.length,isStepFile:isStep},timestamp:Date.now(),sessionId:'debug-session',runId:'run3',hypothesisId:'F'})}).catch(()=>{});
                    // #endregion
                    
                    // Upload attachment to the draft
                    const result = await AuthService.graphRequest(`/me/messages/${draft.id}/attachments`, {
                        method: 'POST',
                        body: JSON.stringify(attachmentPayload)
                    });
                    
                    // #region agent log
                    fetch('http://127.0.0.1:7248/ingest/c8aaba02-7147-41b9-988d-15ca39db2160',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({location:'email-operations.js:820',message:'Attachment upload successful',data:{draftId:draft.id,attachmentName:attachment.name,result:result,isStepFile:isStep},timestamp:Date.now(),sessionId:'debug-session',runId:'run3',hypothesisId:'F'})}).catch(()=>{});
                    // #endregion
                    
                    if (isStep) {
                        console.log(`🔧 [STEP FILE] ✓ Successfully uploaded ${attachment.name} to draft ${draft.id}`);
                        stepFilesUploaded++;
                    } else {
                        console.log(`[saveDraft] ✓ Successfully attached ${attachment.name} to draft ${draft.id}`);
                    }
                    successfulUploads++;
                    
                    // Small delay between attachments to avoid rate limiting
                    if (i < options.attachments.length - 1) {
                        await new Promise(resolve => setTimeout(resolve, 200));
                    }
                } catch (attachmentError) {
                    failedUploads++;
                    if (isStep) {
                        stepFilesFailed++;
                        console.error(`🔧 [STEP FILE] CRITICAL: Failed to upload ${attachment.name}:`, attachmentError);
                        console.error(`  Error type: ${attachmentError.name}, Message: ${attachmentError.message}`);
                    } else {
                        console.error(`[saveDraft] ✗ CRITICAL: Failed to attach ${attachment.name}:`, attachmentError);
                    }
                    // #region agent log
                    fetch('http://127.0.0.1:7248/ingest/c8aaba02-7147-41b9-988d-15ca39db2160',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({location:'email-operations.js:827',message:'Attachment upload failed',data:{draftId:draft.id,attachmentName:attachment.name,errorMessage:attachmentError.message,errorName:attachmentError.name,errorStack:attachmentError.stack,hasContentBytes:!!attachment.contentBytes,contentBytesLength:attachment.contentBytes?.length,isStepFile:isStep},timestamp:Date.now(),sessionId:'debug-session',runId:'run3',hypothesisId:'F'})}).catch(()=>{});
                    // #endregion
                    console.error('[saveDraft] Error details:', {
                        message: attachmentError.message,
                        stack: attachmentError.stack,
                        attachmentName: attachment.name,
                        hasContentBytes: !!attachment.contentBytes,
                        contentBytesLength: attachment.contentBytes ? attachment.contentBytes.length : 0,
                        draftId: draft.id
                    });
                    // Continue with other attachments even if one fails
                }
            }
            
            // STEP file upload summary
            if (stepFilesToUpload.length > 0) {
                console.group('🔧 STEP File Upload Summary');
                console.log(`Total STEP files to upload: ${stepFilesToUpload.length}`);
                console.log(`Successfully uploaded: ${stepFilesUploaded}`);
                console.log(`Failed: ${stepFilesFailed}`);
                if (stepFilesFailed > 0) {
                    console.error(`⚠️ WARNING: ${stepFilesFailed} STEP file(s) failed to upload!`);
                } else if (stepFilesUploaded === stepFilesToUpload.length) {
                    console.log(`✓ All ${stepFilesUploaded} STEP file(s) uploaded successfully`);
                }
                console.groupEnd();
            }
            
            // Check if all uploads failed
            if (successfulUploads === 0 && options.attachments.length > 0) {
                console.error(`[saveDraft] CRITICAL: All ${options.attachments.length} attachment upload(s) failed!`);
                console.error(`[saveDraft] Failed uploads: ${failedUploads}, Successful: ${successfulUploads}`);
                if (stepFilesToUpload.length > 0) {
                    console.error(`[saveDraft] STEP files: ${stepFilesFailed} failed out of ${stepFilesToUpload.length}`);
                }
                // Don't throw - allow draft to be created without attachments, but log critical error
            } else if (successfulUploads > 0 && failedUploads > 0) {
                console.warn(`[saveDraft] Partial success: ${successfulUploads} succeeded, ${failedUploads} failed out of ${options.attachments.length} total`);
                if (stepFilesToUpload.length > 0) {
                    console.warn(`[saveDraft] STEP files: ${stepFilesUploaded} succeeded, ${stepFilesFailed} failed out of ${stepFilesToUpload.length}`);
                }
            } else if (successfulUploads === options.attachments.length) {
                console.log(`[saveDraft] ✓ All ${successfulUploads} attachment(s) uploaded successfully`);
                if (stepFilesToUpload.length > 0) {
                    console.log(`[saveDraft] ✓ All ${stepFilesUploaded} STEP file(s) uploaded successfully`);
                }
            }
            
            // Verify attachments were added by fetching the draft
            let verifiedAttachmentCount = 0;
            let stepFilesInDraft = [];
            let verificationError = null;
            
            try {
                const verifyDraft = await AuthService.graphRequest(`/me/messages/${draft.id}?$expand=attachments&$select=id,subject,attachments`);
                verifiedAttachmentCount = verifyDraft.attachments ? verifyDraft.attachments.length : 0;
                console.log(`✓ Draft ${draft.id} now has ${verifiedAttachmentCount} attachment(s)`);
                
                // Check for STEP files in verified draft
                stepFilesInDraft = (verifyDraft.attachments || []).filter(a => {
                    const ext = (a.name || '').split('.').pop().toLowerCase();
                    return ext === 'step' || ext === 'stp';
                });
                
                if (stepFilesToUpload.length > 0) {
                    console.log(`🔧 STEP files in draft: ${stepFilesInDraft.length} out of ${stepFilesToUpload.length} expected`);
                    if (stepFilesInDraft.length === stepFilesToUpload.length) {
                        console.log(`✓ All STEP files verified in draft:`, stepFilesInDraft.map(a => a.name));
                    } else {
                        console.warn(`⚠️ STEP file verification mismatch: Expected ${stepFilesToUpload.length}, found ${stepFilesInDraft.length}`);
                        console.warn(`  Expected:`, stepFilesToUpload.map(a => a.name));
                        console.warn(`  Found in draft:`, stepFilesInDraft.map(a => a.name));
                    }
                }
                
                // #region agent log
                fetch('http://127.0.0.1:7248/ingest/c8aaba02-7147-41b9-988d-15ca39db2160',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({location:'email-operations.js:844',message:'Attachment verification result',data:{draftId:draft.id,expectedCount:options.attachments.length,actualCount:verifiedAttachmentCount,stepFilesExpected:stepFilesToUpload.length,stepFilesFound:stepFilesInDraft.length,attachments:verifyDraft.attachments?.map(a=>({name:a.name,size:a.size}))},timestamp:Date.now(),sessionId:'debug-session',runId:'run3',hypothesisId:'F'})}).catch(()=>{});
                // #endregion
                
                if (verifiedAttachmentCount === 0) {
                    console.warn('⚠ WARNING: Draft was created but no attachments were found. This may indicate an upload failure.');
                }
            } catch (verifyError) {
                verificationError = verifyError;
                // #region agent log
                fetch('http://127.0.0.1:7248/ingest/c8aaba02-7147-41b9-988d-15ca39db2160',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({location:'email-operations.js:850',message:'Attachment verification failed',data:{draftId:draft.id,errorMessage:verifyError.message},timestamp:Date.now(),sessionId:'debug-session',runId:'run3',hypothesisId:'F'})}).catch(()=>{});
                // #endregion
                console.warn('Could not verify attachments:', verifyError);
            }
            
            // Return draft with upload status
            return {
                draft: draft,
                uploadStatus: {
                    totalAttachments: options.attachments.length,
                    successfulUploads: successfulUploads,
                    failedUploads: failedUploads,
                    stepFilesToUpload: stepFilesToUpload.length,
                    stepFilesUploaded: stepFilesUploaded,
                    stepFilesFailed: stepFilesFailed,
                    verifiedAttachmentCount: verifiedAttachmentCount,
                    stepFilesInDraft: stepFilesInDraft.map(a => a.name),
                    verificationError: verificationError ? verificationError.message : null
                }
            };
        }

        // No attachments - return draft with minimal status
        return {
            draft: draft,
            uploadStatus: {
                totalAttachments: 0,
                successfulUploads: 0,
                failedUploads: 0,
                stepFilesToUpload: 0,
                stepFilesUploaded: 0,
                stepFilesFailed: 0,
                verifiedAttachmentCount: 0,
                stepFilesInDraft: [],
                verificationError: null
            }
        };
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
    /**
     * Format a value for display in email body
     * Handles nested objects and arrays properly
     */
    formatValueForDisplay(value, indent = '') {
        if (value === null || value === undefined) {
            return '';
        }
        
        if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') {
            return String(value);
        }
        
        if (Array.isArray(value)) {
            return value.map(item => this.formatValueForDisplay(item)).join(', ');
        }
        
        if (typeof value === 'object') {
            // Format nested object as indented key-value pairs
            const lines = [];
            for (const [k, v] of Object.entries(value)) {
                const formattedKey = k.replace(/_/g, ' ').replace(/\b\w/g, c => c.toUpperCase());
                const formattedValue = this.formatValueForDisplay(v, indent + '  ');
                if (typeof v === 'object' && v !== null && !Array.isArray(v)) {
                    lines.push(`${indent}  ${formattedKey}:`);
                    lines.push(`${formattedValue}`);
                } else {
                    lines.push(`${indent}  ${formattedKey}: ${formattedValue}`);
                }
            }
            return lines.join('\n');
        }
        
        return String(value);
    },

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
                            const formattedKey = key.replace(/_/g, ' ');
                            const formattedValue = this.formatValueForDisplay(value);
                            if (typeof value === 'object' && value !== null && !Array.isArray(value)) {
                                // Nested object - format on multiple lines
                                text += `- ${formattedKey}:\n${formattedValue}\n`;
                            } else {
                                text += `- ${formattedKey}: ${formattedValue}\n`;
                            }
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
                const jsonStr = JSON.stringify(body, null, 2);
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

/**
 * Folder Category Service
 * Manages email categories/tags based on folder location
 * When an email is moved to a folder, it automatically gets tagged with the folder's category
 */
const FolderCategoryService = {
    // Track if categories have been initialized this session
    categoriesInitialized: false,
    
    // Cache to avoid redundant PATCH requests
    // Maps emailId -> categoryName
    messageCategoryCache: {},

    /**
     * Get array of all folder category names
     * Used to identify which categories belong to this system
     */
    getFolderCategoryNames() {
        if (!Config.FOLDER_CATEGORIES) return [];
        return Object.values(Config.FOLDER_CATEGORIES).map(cat => cat.name);
    },

    /**
     * Get the category definition for a folder name
     * @param {string} folderName - The folder name (e.g., "Sent RFQs", "Quotes")
     * @returns {object|null} Category definition with name and color, or null if not found
     */
    getCategoryForFolder(folderName) {
        if (!Config.FOLDER_CATEGORIES || !folderName) return null;
        return Config.FOLDER_CATEGORIES[folderName] || null;
    },

    /**
     * Ensure all folder categories exist in the user's master category list
     * Should be called once after sign-in
     * @param {boolean} force - If true, bypasses the initialization check and re-runs
     */
    async ensureFolderCategoriesExist(force = false) {
        if (this.categoriesInitialized && !force) {
            console.log('Folder categories already initialized this session');
            return;
        }

        if (!AuthService.isSignedIn()) {
            console.log('Cannot initialize categories - user not signed in');
            return;
        }

        try {
            console.log('Initializing folder categories...');
            
            // Get current master categories
            const response = await AuthService.graphRequest('/me/outlook/masterCategories');
            const existingCategories = response.value || [];
            
            // Create a map of existing categories by name (case-insensitive)
            const existingByName = {};
            existingCategories.forEach(cat => {
                if (cat.displayName) {
                    existingByName[cat.displayName.toLowerCase()] = cat;
                }
            });

            // Ensure each folder category exists with correct color
            for (const folderName in Config.FOLDER_CATEGORIES) {
                const catDef = Config.FOLDER_CATEGORIES[folderName];
                const existingCat = existingByName[catDef.name.toLowerCase()];

                if (!existingCat) {
                    // Create new category
                    try {
                        await AuthService.graphRequest('/me/outlook/masterCategories', {
                            method: 'POST',
                            body: JSON.stringify({
                                displayName: catDef.name,
                                color: catDef.color
                            })
                        });
                        console.log(`Created folder category: "${catDef.name}" (${catDef.color})`);
                    } catch (createError) {
                        // Category might already exist with different casing
                        console.warn(`Could not create category "${catDef.name}":`, createError.message);
                    }
                } else if (existingCat.color !== catDef.color) {
                    // Update color only if different
                    try {
                        await AuthService.graphRequest(
                            `/me/outlook/masterCategories/${encodeURIComponent(existingCat.id)}`,
                            {
                                method: 'PATCH',
                                body: JSON.stringify({ color: catDef.color })
                            }
                        );
                        console.log(`Updated color for category: "${catDef.name}" from ${existingCat.color} to ${catDef.color}`);
                    } catch (updateError) {
                        console.warn(`Could not update category "${catDef.name}":`, updateError.message);
                    }
                } else {
                    console.log(`Category "${catDef.name}" already has correct color (${catDef.color})`);
                }
            }

            this.categoriesInitialized = true;
            console.log('Folder categories initialized successfully');
        } catch (error) {
            console.error('Error initializing folder categories:', error);
        }
    },

    /**
     * Ensure a master category exists with the correct color
     * Creates the category if it doesn't exist, updates the color if it does
     * @param {string} categoryName - The category display name
     * @param {string} color - The color preset (e.g., 'Preset4')
     */
    async ensureCategoryWithColor(categoryName, color) {
        if (!AuthService.isSignedIn()) return;

        try {
            // Get all master categories
            const response = await AuthService.graphRequest('/me/outlook/masterCategories');
            const existingCategories = response.value || [];
            
            // Find existing category (case-insensitive)
            const existing = existingCategories.find(cat => 
                cat.displayName && cat.displayName.toLowerCase() === categoryName.toLowerCase()
            );

            if (!existing) {
                // Create new category with color
                await AuthService.graphRequest('/me/outlook/masterCategories', {
                    method: 'POST',
                    body: JSON.stringify({
                        displayName: categoryName,
                        color: color
                    })
                });
                console.log(`Created master category "${categoryName}" with color ${color}`);
            } else if (existing.color !== color) {
                // Update color if different
                await AuthService.graphRequest(
                    `/me/outlook/masterCategories/${encodeURIComponent(existing.id)}`,
                    {
                        method: 'PATCH',
                        body: JSON.stringify({ color: color })
                    }
                );
                console.log(`Updated master category "${categoryName}" color to ${color}`);
            }
        } catch (error) {
            console.warn(`Could not ensure category "${categoryName}" with color:`, error.message);
        }
    },

    /**
     * Set folder category on an email
     * Removes any existing folder categories and applies the new one
     * Also ensures the master category has the correct color
     * @param {string} emailId - The email ID (Graph API format)
     * @param {string} folderName - The folder name to tag the email with
     */
    async setFolderCategory(emailId, folderName) {
        if (!emailId || !folderName) {
            console.warn('setFolderCategory: Missing emailId or folderName');
            return;
        }

        if (!AuthService.isSignedIn()) {
            console.warn('setFolderCategory: User not signed in');
            return;
        }

        const catDef = this.getCategoryForFolder(folderName);
        if (!catDef) {
            console.log(`No category defined for folder: "${folderName}"`);
            return;
        }

        // Check cache to avoid redundant updates
        if (this.messageCategoryCache[emailId] === catDef.name) {
            console.log(`Email ${emailId} already has category "${catDef.name}" (cached)`);
            return;
        }

        try {
            // IMPORTANT: First ensure the master category exists with correct color
            // This is critical - without this, categories appear without colors
            await this.ensureCategoryWithColor(catDef.name, catDef.color);

            // Get current categories on the email
            const email = await AuthService.graphRequest(
                `/me/messages/${encodeURIComponent(emailId)}?$select=id,categories`
            );
            const currentCategories = (email && email.categories) || [];

            // Get all folder category names to filter out
            const folderCategoryNames = this.getFolderCategoryNames();

            // Remove all folder categories, keep other categories (user's tags, etc.)
            const filteredCategories = currentCategories.filter(cat => {
                return !folderCategoryNames.includes(cat);
            });

            // Add the new folder category
            filteredCategories.push(catDef.name);

            // Update the email
            await AuthService.graphRequest(
                `/me/messages/${encodeURIComponent(emailId)}`,
                {
                    method: 'PATCH',
                    body: JSON.stringify({ categories: filteredCategories })
                }
            );

            // Update cache
            this.messageCategoryCache[emailId] = catDef.name;
            
            console.log(`Set folder category "${catDef.name}" (${catDef.color}) on email ${emailId}`);
        } catch (error) {
            console.error(`Error setting folder category on email ${emailId}:`, error);
            // Don't throw - category failure shouldn't break email operations
        }
    },

    /**
     * Remove all folder categories from an email
     * Preserves user-applied and other categories
     * @param {string} emailId - The email ID (Graph API format)
     */
    async removeFolderCategories(emailId) {
        if (!emailId || !AuthService.isSignedIn()) return;

        try {
            // Get current categories
            const email = await AuthService.graphRequest(
                `/me/messages/${encodeURIComponent(emailId)}?$select=id,categories`
            );
            const currentCategories = (email && email.categories) || [];

            // Get folder category names to remove
            const folderCategoryNames = this.getFolderCategoryNames();

            // Filter out folder categories
            const filteredCategories = currentCategories.filter(cat => {
                return !folderCategoryNames.includes(cat);
            });

            // Only update if something changed
            if (filteredCategories.length !== currentCategories.length) {
                await AuthService.graphRequest(
                    `/me/messages/${encodeURIComponent(emailId)}`,
                    {
                        method: 'PATCH',
                        body: JSON.stringify({ categories: filteredCategories })
                    }
                );
                console.log(`Removed folder categories from email ${emailId}`);
            }

            // Clear cache
            delete this.messageCategoryCache[emailId];
        } catch (error) {
            console.error(`Error removing folder categories from email ${emailId}:`, error);
        }
    },

    /**
     * Clear the message category cache
     * Call when user signs out or as needed
     * @param {string} emailId - Optional specific email ID to clear, or clears all if omitted
     */
    clearCache(emailId) {
        if (emailId) {
            delete this.messageCategoryCache[emailId];
        } else {
            this.messageCategoryCache = {};
            this.categoriesInitialized = false;
        }
    }
};
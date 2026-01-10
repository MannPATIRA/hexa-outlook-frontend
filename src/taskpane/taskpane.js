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
    availableRfqs: []
};

// ==================== INITIALIZATION ====================
Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        console.log('Office.js is ready in Outlook');
        initializeApp();
    } else {
        console.log('Running outside of Outlook - limited functionality');
        initializeApp();
    }
});

function initializeApp() {
    // Load saved settings
    Config.loadSettings();
    
    // Initialize authentication
    initializeAuth();
    
    // Set up event listeners
    setupEventListeners();
    
    // Load initial data
    loadInitialData();
    
    // Update UI based on context
    updateContextUI();
    
    console.log('Procurement Workflow Add-in initialized');
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
    try {
        Helpers.showLoading('Loading open PRs...');
        await loadOpenPRs();
    } catch (error) {
        Helpers.showError('Failed to load initial data: ' + error.message);
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
                await EmailOperations.sendEmail({
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
            Helpers.showSuccess(`All ${successCount} RFQ(s) sent successfully`);
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
    
    Helpers.showElement(document.getElementById('settings-modal'));
}

function closeSettingsModal() {
    Helpers.hideElement(document.getElementById('settings-modal'));
}

function saveSettings() {
    const settings = {
        apiUrl: document.getElementById('api-url').value,
        engineeringEmail: document.getElementById('engineering-email').value,
        autoClassify: document.getElementById('auto-classify').checked,
        autoCreateFolders: document.getElementById('auto-create-folders').checked
    };
    
    Config.saveSettings(settings);
    closeSettingsModal();
    Helpers.showSuccess('Settings saved');
}

// Make functions available globally for inline event handlers
window.previewRFQ = previewRFQ;
window.createSingleDraft = createSingleDraft;
window.handleSupplierCheckboxChange = handleSupplierCheckboxChange;
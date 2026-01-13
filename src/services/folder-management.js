/**
 * Folder Management Service
 * Handles email folder creation and organization using Microsoft Graph API
 */
const FolderManagement = {
    // Cache for folder IDs to avoid repeated lookups
    folderCache: new Map(),

    /**
     * Initialize folder structure for a material
     * Creates the folder hierarchy if it doesn't exist
     */
    async initializeMaterialFolders(materialCode) {
        if (!AuthService.isSignedIn()) {
            throw new Error('Please sign in to manage folders');
        }

        const rootFolderName = materialCode; // e.g., "MAT-12345"
        
        try {
            // Create root folder under Inbox
            const rootFolder = await this.createFolderIfNotExists(rootFolderName);

            // Create subfolders (all at the same level, not nested)
            const subfolders = [
                Config.FOLDERS.SENT_RFQS,
                Config.FOLDERS.QUOTES,
                Config.FOLDERS.CLARIFICATION_REQUESTS,
                Config.FOLDERS.AWAITING_CLARIFICATION,  // Separate folder, not nested
                Config.FOLDERS.AWAITING_ENGINEER,
                Config.FOLDERS.ENGINEER_RESPONSE
            ];

            for (const subfolder of subfolders) {
                await this.createFolderIfNotExists(subfolder, rootFolder.id);
            }

            console.log(`Folder structure created for ${materialCode}`);
            return rootFolder;
        } catch (error) {
            console.error('Error creating folder structure:', error);
            throw error;
        }
    },

    /**
     * Get folder by name within a parent folder
     */
    async getFolderByName(folderName, parentFolderId = null) {
        const cacheKey = `${parentFolderId || 'root'}/${folderName}`;
        
        if (this.folderCache.has(cacheKey)) {
            return this.folderCache.get(cacheKey);
        }

        try {
            const endpoint = parentFolderId 
                ? `/me/mailFolders/${parentFolderId}/childFolders`
                : '/me/mailFolders';
            
            const response = await AuthService.graphRequest(
                `${endpoint}?$filter=displayName eq '${encodeURIComponent(folderName)}'`
            );

            const folder = response.value && response.value[0];
            if (folder) {
                this.folderCache.set(cacheKey, folder);
            }
            return folder || null;
        } catch (error) {
            console.error(`Error finding folder ${folderName}:`, error);
            return null;
        }
    },

    /**
     * Create a folder if it doesn't already exist
     */
    async createFolderIfNotExists(folderName, parentFolderId = null) {
        // Check if folder exists
        let folder = await this.getFolderByName(folderName, parentFolderId);
        
        if (folder) {
            console.log(`Folder already exists: ${folderName}`);
            return folder;
        }

        // Create the folder
        try {
            const endpoint = parentFolderId 
                ? `/me/mailFolders/${parentFolderId}/childFolders`
                : '/me/mailFolders';

            folder = await AuthService.graphRequest(endpoint, {
                method: 'POST',
                body: JSON.stringify({ displayName: folderName })
            });

            // Cache the new folder
            const cacheKey = `${parentFolderId || 'root'}/${folderName}`;
            this.folderCache.set(cacheKey, folder);
            
            console.log(`Created folder: ${folderName}`);
            return folder;
        } catch (error) {
            console.error(`Error creating folder ${folderName}:`, error);
            throw error;
        }
    },

    /**
     * Get folder ID by path (e.g., "MAT-12345/Quotes")
     */
    async getFolderIdByPath(folderPath) {
        const parts = folderPath.split('/');
        let parentId = null;

        for (const part of parts) {
            const folder = await this.getFolderByName(part, parentId);
            if (!folder) {
                return null;
            }
            parentId = folder.id;
        }

        return parentId;
    },

    /**
     * Move an email to a specific folder
     */
    async moveEmailToFolder(emailId, folderPath) {
        if (!AuthService.isSignedIn()) {
            throw new Error('Please sign in to move emails');
        }

        try {
            console.log(`Attempting to move email ${emailId} to folder path: ${folderPath}`);
            
            // Get the destination folder ID
            const folderId = await this.getFolderIdByPath(folderPath);
            
            if (!folderId) {
                console.error(`Folder not found: ${folderPath}`);
                console.error('Available folders may need to be created first');
                throw new Error(`Folder not found: ${folderPath}`);
            }

            console.log(`Found folder ID: ${folderId} for path: ${folderPath}`);

            // Move the message
            const result = await AuthService.graphRequest(`/me/messages/${emailId}/move`, {
                method: 'POST',
                body: JSON.stringify({ destinationId: folderId })
            });

            console.log(`Successfully moved email ${emailId} to ${folderPath} (folder ID: ${folderId})`);
            return result;
        } catch (error) {
            console.error('Error moving email:', error);
            console.error('Email ID:', emailId);
            console.error('Folder path:', folderPath);
            console.error('Error details:', error.message);
            throw error;
        }
    },

    /**
     * Copy an email to a specific folder
     */
    async copyEmailToFolder(emailId, folderPath) {
        if (!AuthService.isSignedIn()) {
            throw new Error('Please sign in to copy emails');
        }

        try {
            const folderId = await this.getFolderIdByPath(folderPath);
            
            if (!folderId) {
                throw new Error(`Folder not found: ${folderPath}`);
            }

            const result = await AuthService.graphRequest(`/me/messages/${emailId}/copy`, {
                method: 'POST',
                body: JSON.stringify({ destinationId: folderId })
            });

            console.log(`Copied email ${emailId} to ${folderPath}`);
            return result;
        } catch (error) {
            console.error('Error copying email:', error);
            throw error;
        }
    },

    /**
     * Get the target folder path based on email classification
     */
    getFolderForClassification(materialCode, classification, subClassification = null) {
        const basePath = materialCode;
        
        switch (classification) {
            case 'quote':
                return `${basePath}/${Config.FOLDERS.QUOTES}`;
            
            case 'clarification_request':
                if (subClassification === 'engineering') {
                    return `${basePath}/${Config.FOLDERS.AWAITING_ENGINEER}`;
                }
                return `${basePath}/${Config.FOLDERS.CLARIFICATION_REQUESTS}`;
            
            case 'engineer_response':
                return `${basePath}/${Config.FOLDERS.ENGINEER_RESPONSE}`;
            
            case 'sent_rfq':
                return `${basePath}/${Config.FOLDERS.SENT_RFQS}`;
            
            default:
                return basePath;
        }
    },

    /**
     * Organize email based on its classification
     */
    async organizeEmail(emailId, materialCode, classification, subClassification = null) {
        // Get the appropriate folder
        const folderPath = this.getFolderForClassification(
            materialCode, 
            classification, 
            subClassification
        );

        // Initialize folders if needed
        const autoCreateFolders = Config.getSetting(
            Config.STORAGE_KEYS.AUTO_CREATE_FOLDERS, 
            true
        );

        if (autoCreateFolders) {
            await this.initializeMaterialFolders(materialCode);
        }

        // Move the email
        await this.moveEmailToFolder(emailId, folderPath);

        return folderPath;
    },

    /**
     * List all subfolders of a material folder
     */
    async listMaterialFolders(materialCode) {
        if (!AuthService.isSignedIn()) {
            throw new Error('Please sign in to list folders');
        }

        try {
            const rootFolder = await this.getFolderByName(materialCode);
            if (!rootFolder) {
                return [];
            }

            const response = await AuthService.graphRequest(
                `/me/mailFolders/${rootFolder.id}/childFolders`
            );

            return response.value || [];
        } catch (error) {
            console.error('Error listing folders:', error);
            return [];
        }
    },

    /**
     * Get emails in a specific folder
     */
    async getEmailsInFolder(folderPath, options = {}) {
        if (!AuthService.isSignedIn()) {
            throw new Error('Please sign in to get emails');
        }

        try {
            const folderId = await this.getFolderIdByPath(folderPath);
            if (!folderId) {
                return [];
            }

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

            const response = await AuthService.graphRequest(endpoint);
            return response.value || [];
        } catch (error) {
            console.error('Error getting emails:', error);
            return [];
        }
    },

    /**
     * Clear the folder cache
     */
    clearCache() {
        this.folderCache.clear();
    },

    /**
     * Get folder path for an email by its folder ID
     * Returns the full path like "MAT-12345/Quotes"
     */
    async getFolderPath(folderId) {
        if (!AuthService.isSignedIn() || !folderId) {
            return '';
        }

        try {
            const pathParts = [];
            let currentFolderId = folderId;
            let maxDepth = 5; // Prevent infinite loops

            while (currentFolderId && maxDepth > 0) {
                const folder = await AuthService.graphRequest(`/me/mailFolders/${currentFolderId}?$select=displayName,parentFolderId`);
                
                if (!folder) break;
                
                pathParts.unshift(folder.displayName);
                
                // Stop if we've reached the root (Inbox, etc.)
                if (!folder.parentFolderId || folder.displayName === 'Inbox') {
                    break;
                }
                
                currentFolderId = folder.parentFolderId;
                maxDepth--;
            }

            return pathParts.join('/');
        } catch (error) {
            console.error('Error getting folder path:', error);
            return '';
        }
    },

    /**
     * Get folder info by ID
     */
    async getFolderById(folderId) {
        if (!AuthService.isSignedIn() || !folderId) {
            return null;
        }

        try {
            return await AuthService.graphRequest(`/me/mailFolders/${folderId}?$select=id,displayName,parentFolderId`);
        } catch (error) {
            console.error('Error getting folder by ID:', error);
            return null;
        }
    },

    /**
     * Get folder structure for a material
     */
    getFolderStructure(materialCode) {
        return {
            root: materialCode,
            subfolders: [
                {
                    name: Config.FOLDERS.SENT_RFQS,
                    path: `${materialCode}/${Config.FOLDERS.SENT_RFQS}`,
                    description: 'Sent RFQ emails'
                },
                {
                    name: Config.FOLDERS.QUOTES,
                    path: `${materialCode}/${Config.FOLDERS.QUOTES}`,
                    description: 'Received quote emails'
                },
                {
                    name: Config.FOLDERS.CLARIFICATION_REQUESTS,
                    path: `${materialCode}/${Config.FOLDERS.CLARIFICATION_REQUESTS}`,
                    description: 'Supplier clarification requests'
                },
                {
                    name: Config.FOLDERS.AWAITING_CLARIFICATION,
                    path: `${materialCode}/${Config.FOLDERS.AWAITING_CLARIFICATION}`,
                    description: 'Awaiting your response'
                },
                {
                    name: Config.FOLDERS.AWAITING_ENGINEER,
                    path: `${materialCode}/${Config.FOLDERS.AWAITING_ENGINEER}`,
                    description: 'Forwarded to engineering team'
                },
                {
                    name: Config.FOLDERS.ENGINEER_RESPONSE,
                    path: `${materialCode}/${Config.FOLDERS.ENGINEER_RESPONSE}`,
                    description: 'Technical responses from engineering'
                }
            ]
        };
    }
};

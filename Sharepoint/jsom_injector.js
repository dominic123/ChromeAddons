// JSOM Injector - Runs in page context to access SharePoint JSOM libraries
// This script is injected into the page and can access the SP object

(function() {
    'use strict';

    console.log('SharePoint Field Creator - JSOM Injector Loaded');

    // Wait for SharePoint libraries to be loaded
    function ensureSPLoaded(callback) {
        if (typeof SP !== 'undefined' && typeof SP.ClientContext !== 'undefined') {
            callback();
        } else {
            // Wait for SP.js to load
            const maxWait = 10000; // 10 seconds max
            const startTime = Date.now();

            const checkInterval = setInterval(() => {
                if (typeof SP !== 'undefined' && typeof SP.ClientContext !== 'undefined') {
                    clearInterval(checkInterval);
                    callback();
                } else if (Date.now() - startTime > maxWait) {
                    clearInterval(checkInterval);
                    console.error('SharePoint libraries failed to load');
                }
            }, 100);
        }
    }

    // Listen for messages from content script
    window.addEventListener('message', function(event) {
        // Only accept messages from same origin
        if (event.origin !== window.location.origin) {
            return;
        }

        const data = event.data;

        if (data.type === 'SP_FIELD_CREATOR_CONNECT') {
            handleConnect(data.siteUrl, data.listName);
        }

        if (data.type === 'SP_FIELD_CREATOR_CREATE_FIELD') {
            handleCreateField(data.siteUrl, data.listName, data.fieldData);
        }

        if (data.type === 'SP_FIELD_CREATOR_CREATE_LIST') {
            handleCreateList(data.siteUrl, data.listName);
        }
    });

    // Handle connection test
    function handleConnect(siteUrl, listName) {
        ensureSPLoaded(function() {
            try {
                const context = SP.ClientContext.get_current();
                const web = context.get_web();

                // Try to get the list by title
                const list = web.get_lists().getByTitle(listName);
                context.load(list, 'Title', 'Id', 'DefaultViewUrl');

                context.executeQueryAsync(
                    function() {
                        sendResponse('SP_FIELD_CREATOR_CONNECT_RESPONSE', {
                            success: true,
                            message: 'Connected successfully!',
                            listInfo: {
                                title: list.get_title(),
                                id: list.get_id(),
                                defaultViewUrl: list.get_defaultViewUrl()
                            }
                        });
                    },
                    function(sender, args) {
                        const errorMessage = args.get_message();
                        console.log('[SP Field Creator] Connection error:', errorMessage);
                        // Check if the error is specifically about list not found
                        // Common error messages: "List 'xxx' does not exist", "Cannot find list", etc.
                        const errorLower = errorMessage.toLowerCase();
                        if (errorLower.indexOf('does not exist') !== -1 ||
                            errorLower.indexOf('cannot find') !== -1 ||
                            errorLower.indexOf('list not found') !== -1 ||
                            errorLower.indexOf('not found') !== -1) {
                            console.log('[SP Field Creator] List not found detected!');
                            sendResponse('SP_FIELD_CREATOR_CONNECT_RESPONSE', {
                                success: false,
                                listNotFound: true,
                                message: `List "${listName}" does not exist. Would you like to create it?`
                            });
                        } else {
                            console.log('[SP Field Creator] Other error - not list not found');
                            sendResponse('SP_FIELD_CREATOR_CONNECT_RESPONSE', {
                                success: false,
                                message: `Error connecting to list: ${errorMessage}`
                            });
                        }
                    }
                );
            } catch (error) {
                sendResponse('SP_FIELD_CREATOR_CONNECT_RESPONSE', {
                    success: false,
                    message: `Error: ${error.message}`
                });
            }
        });
    }

    // Handle field creation
    function handleCreateField(siteUrl, listName, fieldData) {
        console.log('[SP Field Creator] Creating field:', fieldData);
        console.log('[SP Field Creator] List name:', listName);
        console.log('[SP Field Creator] Site URL:', siteUrl);

        ensureSPLoaded(function() {
            try {
                const context = SP.ClientContext.get_current();
                const web = context.get_web();

                console.log('[SP Field Creator] Getting list:', listName);
                const list = web.get_lists().getByTitle(listName);

                // Map data type to SP.FieldType
                const fieldTypeMap = {
                    'text': SP.FieldType.text,
                    'note': SP.FieldType.multiLineText,
                    'number': SP.FieldType.number,
                    'integer': SP.FieldType.integer,
                    'currency': SP.FieldType.currency,
                    'datetime': SP.FieldType.dateTime,
                    'date': SP.FieldType.dateTime,
                    'boolean': SP.FieldType.boolean,
                    'yesno': SP.FieldType.boolean,
                    'user': SP.FieldType.user,
                    'lookup': SP.FieldType.lookup,
                    'choice': SP.FieldType.choice,
                    'url': SP.FieldType.url,
                    'hyperlink': SP.FieldType.url,
                    'counter': SP.FieldType.counter,
                    'calculated': SP.FieldType.calculated,
                    'guid': SP.FieldType.guid
                };

                const fieldType = fieldTypeMap[fieldData.dataType.toLowerCase()] || SP.FieldType.text;
                console.log('[SP Field Creator] Mapped field type:', fieldData.dataType, '->', fieldType);

                // Check if field already exists first
                const existingFields = list.get_fields();
                context.load(existingFields);

                context.executeQueryAsync(
                    function() {
                        console.log('[SP Field Creator] Got existing fields, checking for duplicates');
                        // Check for existing field
                        const fieldEnumerator = existingFields.getEnumerator();
                        let fieldExists = false;
                        let existingFieldNames = [];

                        while (fieldEnumerator.moveNext()) {
                            const existingField = fieldEnumerator.get_current();
                            const fieldName = existingField.get_internalName();
                            existingFieldNames.push(fieldName);
                            if (fieldName === fieldData.fieldName) {
                                fieldExists = true;
                                console.log('[SP Field Creator] Field already exists:', fieldName);
                                break;
                            }
                        }

                        console.log('[SP Field Creator] Existing fields:', existingFieldNames);

                        if (fieldExists) {
                            sendResponse('SP_FIELD_CREATOR_CREATE_FIELD_RESPONSE', {
                                success: false,
                                message: `Field "${fieldData.fieldName}" already exists in the list`
                            });
                            return;
                        }

                        // Add the field to the list using addField methods instead of FieldCreationInformation
                        console.log('[SP Field Creator] Creating new field with type:', fieldType);

                        // Use the simpler SP.FieldCollection.add() method
                        const fields = list.get_fields();
                        let field;

                        // Create field based on type
                        switch(fieldType) {
                            case SP.FieldType.text:
                                field = fields.addFieldAsXml(
                                    `<Field Type="Text" Name="${fieldData.fieldName}" DisplayName="${fieldData.displayName}" ${fieldData.required ? 'Required="TRUE"' : ''} ${fieldData.description ? `Description="${fieldData.description}"` : ''} />`,
                                    true,
                                    SP.AddFieldOptions.addToDefaultContentType
                                );
                                break;
                            case SP.FieldType.multiLineText:
                                field = fields.addFieldAsXml(
                                    `<Field Type="Note" Name="${fieldData.fieldName}" DisplayName="${fieldData.displayName}" NumLines="6" RichText="FALSE" ${fieldData.description ? `Description="${fieldData.description}"` : ''} />`,
                                    true,
                                    SP.AddFieldOptions.addToDefaultContentType
                                );
                                break;
                            case SP.FieldType.number:
                                field = fields.addFieldAsXml(
                                    `<Field Type="Number" Name="${fieldData.fieldName}" DisplayName="${fieldData.displayName}" ${fieldData.required ? 'Required="TRUE"' : ''} ${fieldData.description ? `Description="${fieldData.description}"` : ''} />`,
                                    true,
                                    SP.AddFieldOptions.addToDefaultContentType
                                );
                                break;
                            case SP.FieldType.integer:
                                field = fields.addFieldAsXml(
                                    `<Field Type="Integer" Name="${fieldData.fieldName}" DisplayName="${fieldData.displayName}" ${fieldData.required ? 'Required="TRUE"' : ''} ${fieldData.description ? `Description="${fieldData.description}"` : ''} />`,
                                    true,
                                    SP.AddFieldOptions.addToDefaultContentType
                                );
                                break;
                            case SP.FieldType.currency:
                                field = fields.addFieldAsXml(
                                    `<Field Type="Currency" Name="${fieldData.fieldName}" DisplayName="${fieldData.displayName}" ${fieldData.required ? 'Required="TRUE"' : ''} ${fieldData.description ? `Description="${fieldData.description}"` : ''} />`,
                                    true,
                                    SP.AddFieldOptions.addToDefaultContentType
                                );
                                break;
                            case SP.FieldType.dateTime:
                                field = fields.addFieldAsXml(
                                    `<Field Type="DateTime" Name="${fieldData.fieldName}" DisplayName="${fieldData.displayName}" Format="DateTime" ${fieldData.required ? 'Required="TRUE"' : ''} ${fieldData.description ? `Description="${fieldData.description}"` : ''} />`,
                                    true,
                                    SP.AddFieldOptions.addToDefaultContentType
                                );
                                break;
                            case SP.FieldType.boolean:
                                field = fields.addFieldAsXml(
                                    `<Field Type="Boolean" Name="${fieldData.fieldName}" DisplayName="${fieldData.displayName}" ${fieldData.description ? `Description="${fieldData.description}"` : ''} />`,
                                    true,
                                    SP.AddFieldOptions.addToDefaultContentType
                                );
                                break;
                            case SP.FieldType.user:
                                field = fields.addFieldAsXml(
                                    `<Field Type="User" Name="${fieldData.fieldName}" DisplayName="${fieldData.displayName}" List="UserInfo" ${fieldData.required ? 'Required="TRUE"' : ''} ${fieldData.description ? `Description="${fieldData.description}"` : ''} />`,
                                    true,
                                    SP.AddFieldOptions.addToDefaultContentType
                                );
                                break;
                            case SP.FieldType.choice:
                                field = fields.addFieldAsXml(
                                    `<Field Type="Choice" Name="${fieldData.fieldName}" DisplayName="${fieldData.displayName}" Format="Dropdown" ${fieldData.required ? 'Required="TRUE"' : ''} ${fieldData.description ? `Description="${fieldData.description}"` : ''} />`,
                                    true,
                                    SP.AddFieldOptions.addToDefaultContentType
                                );
                                break;
                            case SP.FieldType.url:
                                field = fields.addFieldAsXml(
                                    `<Field Type="URL" Name="${fieldData.fieldName}" DisplayName="${fieldData.displayName}" ${fieldData.description ? `Description="${fieldData.description}"` : ''} />`,
                                    true,
                                    SP.AddFieldOptions.addToDefaultContentType
                                );
                                break;
                            default:
                                // Default to text for unknown types
                                field = fields.addFieldAsXml(
                                    `<Field Type="Text" Name="${fieldData.fieldName}" DisplayName="${fieldData.displayName}" ${fieldData.required ? 'Required="TRUE"' : ''} ${fieldData.description ? `Description="${fieldData.description}"` : ''} />`,
                                    true,
                                    SP.AddFieldOptions.addToDefaultContentType
                                );
                        }

                        context.load(field);

                        context.executeQueryAsync(
                            function() {
                                console.log('[SP Field Creator] Field created successfully!');
                                sendResponse('SP_FIELD_CREATOR_CREATE_FIELD_RESPONSE', {
                                    success: true,
                                    message: `Field "${fieldData.displayName}" created successfully`,
                                    fieldId: field.get_id()
                                });
                            },
                            function(sender, args) {
                                console.error('[SP Field Creator] Error creating field:', args.get_message());
                                sendResponse('SP_FIELD_CREATOR_CREATE_FIELD_RESPONSE', {
                                    success: false,
                                    message: `Error creating field: ${args.get_message()}`
                                });
                            }
                        );
                    },
                    function(sender, args) {
                        console.error('[SP Field Creator] Error checking existing fields:', args.get_message());
                        sendResponse('SP_FIELD_CREATOR_CREATE_FIELD_RESPONSE', {
                            success: false,
                            message: `Error checking existing fields: ${args.get_message()}`
                        });
                    }
                );
            } catch (error) {
                console.error('[SP Field Creator] Exception:', error);
                sendResponse('SP_FIELD_CREATOR_CREATE_FIELD_RESPONSE', {
                    success: false,
                    message: `Error: ${error.message}\nStack: ${error.stack}`
                });
            }
        });
    }

    // Handle list creation
    function handleCreateList(siteUrl, listName) {
        ensureSPLoaded(function() {
            try {
                const context = SP.ClientContext.get_current();
                const web = context.get_web();

                // Create list creation information
                const listCreationInfo = new SP.ListCreationInformation();
                listCreationInfo.set_title(listName);
                listCreationInfo.set_templateType(SP.ListTemplateType.genericList);
                listCreationInfo.set_quickLaunchOption(SP.QuickLaunchOptions.on);

                // Add the list to the web
                const newList = web.get_lists().add(listCreationInfo);

                context.load(newList, 'Title', 'Id', 'DefaultViewUrl');

                context.executeQueryAsync(
                    function() {
                        console.log('[SP Field Creator] List created successfully!');
                        sendResponse('SP_FIELD_CREATOR_CREATE_LIST_RESPONSE', {
                            success: true,
                            message: `List "${listName}" created successfully`,
                            listInfo: {
                                title: newList.get_title(),
                                id: newList.get_id(),
                                defaultViewUrl: newList.get_defaultViewUrl()
                            }
                        });
                    },
                    function(sender, args) {
                        console.error('[SP Field Creator] Error creating list:', args.get_message());
                        sendResponse('SP_FIELD_CREATOR_CREATE_LIST_RESPONSE', {
                            success: false,
                            message: `Error creating list: ${args.get_message()}`
                        });
                    }
                );
            } catch (error) {
                console.error('[SP Field Creator] Exception creating list:', error);
                sendResponse('SP_FIELD_CREATOR_CREATE_LIST_RESPONSE', {
                    success: false,
                    message: `Error: ${error.message}`
                });
            }
        });
    }

    // Send response back to content script
    function sendResponse(type, response) {
        window.postMessage({
            type: type,
            response: response
        }, window.location.origin);
    }

    // Notify that injector is ready
    console.log('SharePoint Field Creator - JSOM Injector Ready');
})();

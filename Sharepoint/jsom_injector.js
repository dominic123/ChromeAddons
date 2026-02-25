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

        if (data.type === 'SP_FIELD_CREATOR_GET_ALL_LISTS') {
            handleGetAllLists();
        }

        if (data.type === 'SP_FIELD_CREATOR_PREVIEW_ITEMS') {
            handlePreviewItems(data.listTitle, data.camlQuery, data.folderPath);
        }

        if (data.type === 'SP_FIELD_CREATOR_DELETE_LIST') {
            handleDeleteListOperation(data.listTitle);
        }

        if (data.type === 'SP_FIELD_CREATOR_DELETE_ITEMS') {
            handleDeleteItemsOperation(data.listTitle, data.camlQuery, data.folderPath);
        }

        if (data.type === 'SP_FIELD_CREATOR_GET_LIST_FIELDS') {
            handleGetListFields(data.listTitle);
        }

        if (data.type === 'SP_FIELD_CREATOR_FILTER_ITEMS') {
            handleFilterListItems(data.listTitle, data.camlQuery, data.rowLimit);
        }

        if (data.type === 'SP_FIELD_CREATOR_GET_ALL_LISTS_WITH_FIELDS') {
            handleGetAllListsWithFields();
        }

        if (data.type === 'SP_FIELD_CREATOR_SEARCH_ALL_LISTS_BY_FIELD') {
            handleSearchAllListsByField(data.fieldName, data.camlQuery, data.rowLimit);
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

    // Handle get all lists
    function handleGetAllLists() {
        ensureSPLoaded(function() {
            try {
                const context = SP.ClientContext.get_current();
                const web = context.get_web();
                const lists = web.get_lists();

                context.load(lists, 'Include(Title, ItemCount, BaseType)');

                context.executeQueryAsync(
                    function() {
                        console.log('[SP Field Creator] Got all lists successfully!');
                        const listArray = [];
                        const enumerator = lists.getEnumerator();

                        while (enumerator.moveNext()) {
                            const list = enumerator.get_current();
                            listArray.push({
                                title: list.get_title(),
                                itemCount: list.get_itemCount(),
                                baseType: list.get_baseType()
                            });
                        }

                        // Sort by title
                        listArray.sort((a, b) => a.title.localeCompare(b.title));

                        sendResponse('SP_FIELD_CREATOR_GET_ALL_LISTS_RESPONSE', {
                            success: true,
                            lists: listArray
                        });
                    },
                    function(sender, args) {
                        console.error('[SP Field Creator] Error getting lists:', args.get_message());
                        sendResponse('SP_FIELD_CREATOR_GET_ALL_LISTS_RESPONSE', {
                            success: false,
                            message: `Error fetching lists: ${args.get_message()}`
                        });
                    }
                );
            } catch (error) {
                console.error('[SP Field Creator] Exception getting lists:', error);
                sendResponse('SP_FIELD_CREATOR_GET_ALL_LISTS_RESPONSE', {
                    success: false,
                    message: `Error: ${error.message}`
                });
            }
        });
    }

    // Handle preview items
    function handlePreviewItems(listTitle, camlQuery, folderPath) {
        ensureSPLoaded(function() {
            try {
                const context = SP.ClientContext.get_current();
                const web = context.get_web();
                const list = web.get_lists().getByTitle(listTitle);

                const caml = new SP.CamlQuery();
                caml.set_viewXml(camlQuery);

                let items;
                if (folderPath) {
                    caml.set_folderServerRelativeUrl(folderPath);
                    items = list.getItems(caml);
                } else {
                    items = list.getItems(caml);
                }

                context.load(items, 'Include(Id, Title, FileRef, FileSystemObjectType)');

                context.executeQueryAsync(
                    function() {
                        console.log('[SP Field Creator] Got items successfully!');
                        const itemArray = [];
                        const enumerator = items.getEnumerator();
                        while (enumerator.moveNext()) {
                            const item = enumerator.get_current();
                            itemArray.push({
                                id: item.get_id(),
                                title: item.get_item('Title') || '(no title)',
                                fileRef: item.get_item('FileRef') || '',
                                isFolder: item.get_fileSystemObjectType() === 1
                            });
                        }
                        sendResponse('SP_FIELD_CREATOR_PREVIEW_ITEMS_RESPONSE', {
                            success: true,
                            items: itemArray
                        });
                    },
                    function(sender, args) {
                        console.error('[SP Field Creator] Error getting items:', args.get_message());
                        sendResponse('SP_FIELD_CREATOR_PREVIEW_ITEMS_RESPONSE', {
                            success: false,
                            message: `Error fetching items: ${args.get_message()}`
                        });
                    }
                );
            } catch (error) {
                console.error('[SP Field Creator] Exception getting items:', error);
                sendResponse('SP_FIELD_CREATOR_PREVIEW_ITEMS_RESPONSE', {
                    success: false,
                    message: `Error: ${error.message}`
                });
            }
        });
    }

    // Handle delete list
    function handleDeleteListOperation(listTitle) {
        ensureSPLoaded(function() {
            try {
                const context = SP.ClientContext.get_current();
                const web = context.get_web();
                const list = web.get_lists().getByTitle(listTitle);

                list.deleteObject();

                context.executeQueryAsync(
                    function() {
                        console.log('[SP Field Creator] List deleted successfully!');
                        sendResponse('SP_FIELD_CREATOR_DELETE_LIST_RESPONSE', {
                            success: true,
                            message: `List deleted successfully`
                        });
                    },
                    function(sender, args) {
                        console.error('[SP Field Creator] Error deleting list:', args.get_message());
                        sendResponse('SP_FIELD_CREATOR_DELETE_LIST_RESPONSE', {
                            success: false,
                            message: `Error deleting list: ${args.get_message()}`
                        });
                    }
                );
            } catch (error) {
                console.error('[SP Field Creator] Exception deleting list:', error);
                sendResponse('SP_FIELD_CREATOR_DELETE_LIST_RESPONSE', {
                    success: false,
                    message: `Error: ${error.message}`
                });
            }
        });
    }

    // Handle delete items
    function handleDeleteItemsOperation(listTitle, camlQuery, folderPath) {
        ensureSPLoaded(function() {
            try {
                const context = SP.ClientContext.get_current();
                const web = context.get_web();
                const list = web.get_lists().getByTitle(listTitle);

                const caml = new SP.CamlQuery();
                caml.set_viewXml(camlQuery);

                let items;
                if (folderPath) {
                    caml.set_folderServerRelativeUrl(folderPath);
                    items = list.getItems(caml);
                } else {
                    items = list.getItems(caml);
                }

                context.load(items, 'Include(Id)');

                context.executeQueryAsync(
                    function() {
                        console.log('[SP Field Creator] Got items, now deleting...');
                        const itemArray = [];
                        const enumerator = items.getEnumerator();
                        while (enumerator.moveNext()) {
                            const item = enumerator.get_current();
                            itemArray.push(item.get_id());
                        }

                        // Delete items
                        itemArray.forEach(id => {
                            const item = list.getItemById(id);
                            item.deleteObject();
                        });

                        context.executeQueryAsync(
                            function() {
                                console.log('[SP Field Creator] Items deleted successfully!');
                                sendResponse('SP_FIELD_CREATOR_DELETE_ITEMS_RESPONSE', {
                                    success: true,
                                    deleted: itemArray.length,
                                    message: `${itemArray.length} item(s) deleted`
                                });
                            },
                            function(sender, args) {
                                console.error('[SP Field Creator] Error deleting items:', args.get_message());
                                sendResponse('SP_FIELD_CREATOR_DELETE_ITEMS_RESPONSE', {
                                    success: false,
                                    message: `Error deleting items: ${args.get_message()}`
                                });
                            }
                        );
                    },
                    function(sender, args) {
                        console.error('[SP Field Creator] Error getting items to delete:', args.get_message());
                        sendResponse('SP_FIELD_CREATOR_DELETE_ITEMS_RESPONSE', {
                            success: false,
                            message: `Error getting items: ${args.get_message()}`
                        });
                    }
                );
            } catch (error) {
                console.error('[SP Field Creator] Exception deleting items:', error);
                sendResponse('SP_FIELD_CREATOR_DELETE_ITEMS_RESPONSE', {
                    success: false,
                    message: `Error: ${error.message}`
                });
            }
        });
    }

    // Handle get list fields
    function handleGetListFields(listTitle) {
        ensureSPLoaded(function() {
            try {
                const context = SP.ClientContext.get_current();
                const web = context.get_web();
                const list = web.get_lists().getByTitle(listTitle);

                const fields = list.get_fields();
                context.load(fields, 'Include(Title,InternalName,FieldTypeKind,TypeAsString,ReadOnlyField,Hidden,Required)');

                context.executeQueryAsync(
                    function() {
                        const fieldList = [];
                        const fieldEnumerator = fields.getEnumerator();

                        while (fieldEnumerator.moveNext()) {
                            const field = fieldEnumerator.get_current();

                            // Skip hidden system fields
                            if (field.get_hidden() && !field.get_title()) {
                                continue;
                            }

                            // Skip certain system fields
                            const skipFields = ['ContentType', 'ContentTypeId', 'MetaInfo', 'ScopeId',
                                '_Level', '_IsCurrentVersion', 'ItemChildCount', 'FolderChildCount'];
                            if (skipFields.indexOf(field.get_internalName()) !== -1) {
                                continue;
                            }

                            fieldList.push({
                                title: field.get_title() || field.get_internalName(),
                                internalName: field.get_internalName(),
                                type: field.get_typeAsString() || 'Text',
                                fieldTypeKind: field.get_fieldTypeKind(),
                                readOnly: field.get_readOnlyField(),
                                hidden: field.get_hidden(),
                                required: field.get_required()
                            });
                        }

                        // Sort by title
                        fieldList.sort((a, b) => a.title.localeCompare(b.title));

                        sendResponse('SP_FIELD_CREATOR_GET_LIST_FIELDS_RESPONSE', {
                            success: true,
                            fields: fieldList
                        });
                    },
                    function(sender, args) {
                        console.error('[SP Field Creator] Error getting list fields:', args.get_message());
                        sendResponse('SP_FIELD_CREATOR_GET_LIST_FIELDS_RESPONSE', {
                            success: false,
                            message: `Error: ${args.get_message()}`
                        });
                    }
                );
            } catch (error) {
                console.error('[SP Field Creator] Exception getting list fields:', error);
                sendResponse('SP_FIELD_CREATOR_GET_LIST_FIELDS_RESPONSE', {
                    success: false,
                    message: `Error: ${error.message}`
                });
            }
        });
    }

    // Handle filter list items
    function handleFilterListItems(listTitle, camlQuery, rowLimit) {
        ensureSPLoaded(function() {
            try {
                const context = SP.ClientContext.get_current();
                const web = context.get_web();
                const list = web.get_lists().getByTitle(listTitle);

                // Parse CAML query
                const camlQueryObj = new SP.CamlQuery();
                camlQueryObj.set_viewXml(camlQuery);

                const items = list.getItems(camlQueryObj);

                // Get list fields first
                const fields = list.get_fields();
                context.load(fields);
                context.load(items, 'Include(ID)');

                context.executeQueryAsync(
                    function() {
                        // Get field names to load
                        const fieldNames = ['ID', 'Title'];
                        const fieldEnumerator = fields.getEnumerator();
                        while (fieldEnumerator.moveNext()) {
                            const field = fieldEnumerator.get_current();
                            const internalName = field.get_internalName();
                            // Skip certain system fields
                            if (!internalName.startsWith('_') || internalName === '_ModerationStatus') {
                                fieldNames.push(internalName);
                            }
                        }

                        // Now load items with all fields
                        const loadExpressions = ['Include(' + fieldNames.join(', ') + ')'];
                        context.load(items, loadExpressions.join(''));

                        context.executeQueryAsync(
                            function() {
                                const results = [];
                                const itemEnumerator = items.getEnumerator();

                                while (itemEnumerator.moveNext()) {
                                    const item = itemEnumerator.get_current();
                                    const itemData = {};

                                    // Get values for each field
                                    for (var i = 0; i < fieldNames.length; i++) {
                                        var fieldName = fieldNames[i];
                                        try {
                                            var value = item.get_item(fieldName);

                                            if (value === null) {
                                                itemData[fieldName] = '';
                                            } else if (typeof value === 'object') {
                                                // Handle lookup fields
                                                if (value.get_lookupValue) {
                                                    itemData[fieldName] = value.get_lookupValue();
                                                    itemData[fieldName + '_Id'] = value.get_lookupId();
                                                }
                                                // Handle user fields
                                                else if (typeof value.get_title === 'function') {
                                                    itemData[fieldName] = value.get_title();
                                                }
                                                // Handle Date fields
                                                else if (value instanceof Date) {
                                                    itemData[fieldName] = value.toLocaleString();
                                                }
                                                // Handle other objects
                                                else {
                                                    try {
                                                        itemData[fieldName] = JSON.stringify(value);
                                                    } catch (e) {
                                                        itemData[fieldName] = '[Object]';
                                                    }
                                                }
                                            } else {
                                                itemData[fieldName] = String(value);
                                            }
                                        } catch (e) {
                                            // Field doesn't exist or can't be accessed
                                            itemData[fieldName] = '';
                                        }
                                    }

                                    results.push(itemData);
                                }

                                sendResponse('SP_FIELD_CREATOR_FILTER_ITEMS_RESPONSE', {
                                    success: true,
                                    results: results
                                });
                            },
                            function(sender, args) {
                                console.error('[SP Field Creator] Error loading item field values:', args.get_message());
                                sendResponse('SP_FIELD_CREATOR_FILTER_ITEMS_RESPONSE', {
                                    success: false,
                                    message: `Error loading field values: ${args.get_message()}`
                                });
                            }
                        );
                    },
                    function(sender, args) {
                        console.error('[SP Field Creator] Error filtering items:', args.get_message());
                        sendResponse('SP_FIELD_CREATOR_FILTER_ITEMS_RESPONSE', {
                            success: false,
                            message: `Error: ${args.get_message()}`
                        });
                    }
                );
            } catch (error) {
                console.error('[SP Field Creator] Exception filtering items:', error);
                sendResponse('SP_FIELD_CREATOR_FILTER_ITEMS_RESPONSE', {
                    success: false,
                    message: `Error: ${error.message}`
                });
            }
        });
    }

    // Handle get all lists with fields
    function handleGetAllListsWithFields() {
        ensureSPLoaded(function() {
            try {
                var context = SP.ClientContext.get_current();
                var web = context.get_web();
                var lists = web.get_lists();

                context.load(lists, 'Include(Title,Id,ItemCount)');

                context.executeQueryAsync(
                    function() {
                        var allLists = [];
                        var listEnumerator = lists.getEnumerator();

                        while (listEnumerator.moveNext()) {
                            var list = listEnumerator.get_current();
                            allLists.push({
                                title: list.get_title(),
                                id: list.get_id(),
                                itemCount: list.get_itemCount()
                            });
                        }

                        // Now get fields for each list - process sequentially
                        var listsWithFields = [];
                        var processedCount = 0;

                        function processNextList() {
                            if (processedCount >= allLists.length) {
                                // All lists processed
                                sendResponse('SP_FIELD_CREATOR_GET_ALL_LISTS_WITH_FIELDS_RESPONSE', {
                                    success: true,
                                    lists: listsWithFields
                                });
                                return;
                            }

                            var listInfo = allLists[processedCount];
                            var listTitle = listInfo.title;

                            // Load fields for this list
                            var currentList = web.get_lists().getById(listInfo.id);
                            var fields = currentList.get_fields();

                            context.load(fields, 'Include(Title,InternalName,TypeAsString,FieldTypeKind,ReadOnlyField,Hidden)');

                            context.executeQueryAsync(
                                function() {
                                    var fieldList = [];
                                    var fieldEnumerator = fields.getEnumerator();

                                    while (fieldEnumerator.moveNext()) {
                                        var field = fieldEnumerator.get_current();
                                        if (!field.get_hidden()) {
                                            fieldList.push({
                                                title: field.get_title(),
                                                internalName: field.get_internalName(),
                                                type: field.get_typeAsString()
                                            });
                                        }
                                    }

                                    listsWithFields.push({
                                        title: listTitle,
                                        fields: fieldList
                                    });

                                    processedCount++;
                                    setTimeout(processNextList, 50); // Small delay between lists
                                },
                                function(sender, args) {
                                    // Field loading failed, continue with next
                                    console.warn('[SP Field Creator] Could not load fields for list:', listTitle);
                                    listsWithFields.push({
                                        title: listTitle,
                                        fields: []
                                    });

                                    processedCount++;
                                    setTimeout(processNextList, 50);
                                }
                            );
                        }

                        // Start processing
                        processNextList();
                    },
                    function(sender, args) {
                        console.error('[SP Field Creator] Error getting lists:', args.get_message());
                        sendResponse('SP_FIELD_CREATOR_GET_ALL_LISTS_WITH_FIELDS_RESPONSE', {
                            success: false,
                            message: 'Error: ' + args.get_message()
                        });
                    }
                );
            } catch (error) {
                console.error('[SP Field Creator] Exception getting lists with fields:', error);
                sendResponse('SP_FIELD_CREATOR_GET_ALL_LISTS_WITH_FIELDS_RESPONSE', {
                    success: false,
                    message: 'Error: ' + (error.message || error.toString())
                });
            }
        });
    }

    // Handle search all lists by field
    function handleSearchAllListsByField(fieldName, camlQuery, rowLimit) {
        console.log('[SP Field Creator] handleSearchAllListsByField called with fieldName:', fieldName);
        ensureSPLoaded(function() {
            try {
                var context = SP.ClientContext.get_current();
                var web = context.get_web();
                var lists = web.get_lists();

                context.load(lists, 'Include(Title,Id)');

                context.executeQueryAsync(
                    function() {
                        console.log('[SP Field Creator] Got all lists, count:', lists.get_count());
                        var allLists = [];
                        var listEnumerator = lists.getEnumerator();

                        while (listEnumerator.moveNext()) {
                            var list = listEnumerator.get_current();
                            allLists.push({
                                title: list.get_title(),
                                id: list.get_id()
                            });
                        }

                        console.log('[SP Field Creator] Processing', allLists.length, 'lists for field:', fieldName);

                        // Search each list for the field - process sequentially
                        var searchResults = [];
                        var listsWithField = [];
                        var log = [];
                        var totalLists = allLists.length;
                        var processedCount = 0;

                        function processNextList() {
                            if (processedCount >= allLists.length) {
                                // All lists processed
                                console.log('[SP Field Creator] All lists processed. Results:', searchResults.length, 'lists with items');
                                sendResponse('SP_FIELD_CREATOR_SEARCH_ALL_LISTS_BY_FIELD_RESPONSE', {
                                    success: true,
                                    results: searchResults,
                                    listsWithField: listsWithField,
                                    totalLists: totalLists,
                                    fieldName: fieldName,
                                    log: log
                                });
                                return;
                            }

                            var listInfo = allLists[processedCount];
                            var listTitle = listInfo.title;
                            console.log('[SP Field Creator] Processing list:', listTitle, '(', processedCount + 1, 'of', totalLists, ')');

                            // First check if the field exists in this list
                            var currentList = web.get_lists().getById(listInfo.id);
                            var fields = currentList.get_fields();
                            context.load(fields);

                            context.executeQueryAsync(
                                function() {
                                    var fieldExists = false;
                                    var fieldEnumerator = fields.getEnumerator();

                                    while (fieldEnumerator.moveNext()) {
                                        var field = fieldEnumerator.get_current();
                                        if (field.get_internalName() === fieldName || field.get_title() === fieldName) {
                                            fieldExists = true;
                                            break;
                                        }
                                    }

                                    if (fieldExists) {
                                        listsWithField.push(listTitle);
                                        console.log('[SP Field Creator] Field found in:', listTitle);

                                        // Now search for items with the filter
                                        var camlQueryObj = new SP.CamlQuery();
                                        camlQueryObj.set_viewXml(camlQuery);

                                        var items = currentList.getItems(camlQueryObj);
                                        context.load(items);

                                        context.executeQueryAsync(
                                            function() {
                                                var itemList = [];
                                                var itemEnumerator = items.getEnumerator();

                                                while (itemEnumerator.moveNext()) {
                                                    var item = itemEnumerator.get_current();
                                                    var itemData = { ID: item.get_id() };

                                                    // Get Title field
                                                    try {
                                                        itemData['Title'] = item.get_item('Title') || '';
                                                    } catch (e) {
                                                        itemData['Title'] = '';
                                                    }

                                                    // Get the searched field value
                                                    try {
                                                        var fieldValue = item.get_item(fieldName);
                                                        if (fieldValue !== null && typeof fieldValue === 'object') {
                                                            if (typeof fieldValue.get_lookupValue === 'function') {
                                                                itemData[fieldName] = fieldValue.get_lookupValue();
                                                            } else if (typeof fieldValue.get_title === 'function') {
                                                                itemData[fieldName] = fieldValue.get_title();
                                                            } else {
                                                                itemData[fieldName] = '[Object]';
                                                            }
                                                        } else if (fieldValue !== null) {
                                                            itemData[fieldName] = String(fieldValue);
                                                        } else {
                                                            itemData[fieldName] = '';
                                                        }
                                                    } catch (e) {
                                                        itemData[fieldName] = '';
                                                    }

                                                    itemList.push(itemData);
                                                }

                                                console.log('[SP Field Creator] Found', itemList.length, 'items in', listTitle);

                                                if (itemList.length > 0) {
                                                    searchResults.push({
                                                        listTitle: listTitle,
                                                        items: itemList
                                                    });
                                                    log.push({
                                                        status: 'success',
                                                        message: 'List "' + listTitle + '": Found ' + itemList.length + ' items'
                                                    });
                                                } else {
                                                    log.push({
                                                        status: 'success',
                                                        message: 'List "' + listTitle + '": Field found, but no matching items'
                                                    });
                                                }

                                                processedCount++;
                                                setTimeout(processNextList, 50); // Small delay between lists
                                            },
                                            function(sender, args) {
                                                console.error('[SP Field Creator] Error searching items in', listTitle, ':', args.get_message());
                                                log.push({
                                                    status: 'error',
                                                    message: 'List "' + listTitle + '": Error searching items - ' + args.get_message()
                                                });
                                                processedCount++;
                                                setTimeout(processNextList, 50);
                                            }
                                        );
                                    } else {
                                        console.log('[SP Field Creator] Field NOT found in:', listTitle);
                                        log.push({
                                            status: 'skipped',
                                            message: 'List "' + listTitle + '": Field "' + fieldName + '" not found - skipped'
                                        });
                                        processedCount++;
                                        setTimeout(processNextList, 50);
                                    }
                                },
                                function(sender, args) {
                                    console.error('[SP Field Creator] Error checking fields for', listTitle, ':', args.get_message());
                                    log.push({
                                        status: 'error',
                                        message: 'List "' + listTitle + '": Error checking fields - ' + args.get_message()
                                    });
                                    processedCount++;
                                    setTimeout(processNextList, 50);
                                }
                            );
                        }

                        // Start processing
                        processNextList();
                    },
                    function(sender, args) {
                        console.error('[SP Field Creator] Error getting lists:', args.get_message());
                        sendResponse('SP_FIELD_CREATOR_SEARCH_ALL_LISTS_BY_FIELD_RESPONSE', {
                            success: false,
                            message: 'Error: ' + args.get_message()
                        });
                    }
                );
            } catch (error) {
                console.error('[SP Field Creator] Exception searching all lists:', error);
                sendResponse('SP_FIELD_CREATOR_SEARCH_ALL_LISTS_BY_FIELD_RESPONSE', {
                    success: false,
                    message: 'Error: ' + (error.message || error.toString())
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

//Ews request to get item info
function getItemDataRequest(itemId) {
    var soapToGetItemData = '<?xml version="1.0" encoding="utf-8"?>' +
                    '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
                    '               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"' +
                    '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
                    '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
                    '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
                    '  <soap:Header>' +
                    '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
                    '  </soap:Header>' +
                    '  <soap:Body>' +
                    '    <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages"' +
                    '             xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
                    '      <ItemShape>' +
                    '        <t:BaseShape>IdOnly</t:BaseShape>' +
                    '        <t:AdditionalProperties>' +
                    '            <t:FieldURI FieldURI="item:Attachments" /> ' +
                    '        </t:AdditionalProperties> ' +
                    '      </ItemShape>' +
                    '      <ItemIds>' +
                    '        <t:ItemId Id="' + itemId + '"/>' +
                    '      </ItemIds>' +
                    '    </GetItem>' +
                    '  </soap:Body>' +
                    '</soap:Envelope>';

    return soapToGetItemData;
}

//Ews request to send the modified item
function getSendItemRequest(itemId, changeKey) {
    var soapSendItemRequest = '<?xml version="1.0" encoding="utf-8"?>' +
                            '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
                            '               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"' +
                            '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
                            '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
                            '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
                            '  <soap:Header>' +
                            '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
                            '  </soap:Header>' +
                            '  <soap:Body> ' +
                            '    <SendItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages" SaveItemToFolder="true"> ' +
                            '      <ItemIds> ' +
                            '        <t:ItemId Id="' + itemId + '" ChangeKey="' + changeKey + '" /> ' +
                            '      </ItemIds> ' +
                            '      <m:SavedItemFolderId>' +
                            '         <t:DistinguishedFolderId Id="sentitems" />' +
                            '      </m:SavedItemFolderId>' +
                            '    </SendItem> ' +
                            '  </soap:Body> ' +
                            '</soap:Envelope> ';
    return soapSendItemRequest;
}

// t:Content : Contains the Base64-encoded contents of the file attachment.
// https://msdn.microsoft.com/en-us/library/office/aa580492(v=exchg.150).aspx
function getAttachmentXml(fileName, fileData) {
    var attachmentXml = 
        '<t:FileAttachment>'+
            '<t:Name>'+fileName+'</t:Name>'+
            '<t:Content>'+fileData+'</t:Content>'+
        '</t:FileAttachment>';
    return attachmentXml;
}

// 
function getCreateFileAttachmentRequest(itemId, changeKey, attachmentsXml) {
    var soapCreateFileAttachmentRequest = 
        '<?xml version="1.0" encoding="utf-8"?>'+
        '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"'+
                        'xmlns:xsd="http://www.w3.org/2001/XMLSchema"'+
                        'xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"'+
                        'xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">'
            '<soap:Body>'+
                '<CreateAttachment xmlns="http://schemas.microsoft.com/exchange/services/2006/messages"'+
                                  'xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">'+
                    '<ParentItemId Id="'+itemId+'" ChangeKey="'+changeKey+'"/>'+
                    '<Attachments>'+
                        attachmentsXml+
                    '</Attachments>'+
                '</CreateAttachment>'+
            '</soap:Body>'+
        '</soap:Envelope>';
    return soapCreateFileAttachmentRequest;
}

function getAttachmentIdXml(attachmentId) {
    return '<t:AttachmentId Id="'+attachmentId+'"/>';
}

// The DeleteAttachment operation is used to delete file and item attachments from an existing item in the Exchange store.
// note:This operation allows you to delete one or more attachments by ID.
// https://msdn.microsoft.com/en-us/library/office/aa580782(v=exchg.150).aspx
function getDeleteFileAttachmentRequest(attachmentIdsXml) {
    var soapDeleteFileAttachmentRequest = 
        '<?xml version="1.0" encoding="utf-8"?>'+
        '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"'+
                        'xmlns:xsd="http://www.w3.org/2001/XMLSchema"'+
                        'xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"'+
                        'xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">'
            '<soap:Body>'+
                '<DeleteAttachment xmlns="http://schemas.microsoft.com/exchange/services/2006/messages"'+
                                  'xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">'+
                    '<Attachments>'+
                        attachmentIdsXml+
                    '</Attachments>'+
                '</DeleteAttachment>'+
            '</soap:Body>'+
        '</soap:Envelope>';
    return soapDeleteFileAttachmentRequest;
}
// Outlook Service - Extract data from current email with better error handling

export interface OutlookItemData {
  subject: string;
  body: string;
  bodyHtml?: string;
  sender: string;
  senderEmail: string;
  attachments: OutlookAttachment[];
  receivedDate?: Date;
  cc?: string[];
  to?: string[];
  conversationId?: string;
}

export interface OutlookAttachment {
  id: string;
  name: string;
  size?: number;
  contentType?: string;
  isFile?: boolean;
  url?: string;
}

// Get Outlook item data synchronously (basic info)
export const getOutlookItemData = (): OutlookItemData | null => {
  try {
    const item = Office.context.mailbox.item;
    
    if (!item) {
      console.warn('No Outlook item selected');
      return null;
    }

    // Get subject
    let subject = '';
    try {
      subject = item.subject || 'No Subject';
    } catch (e) {
      console.warn('Could not get subject');
    }

    // Get sender
    let sender = '';
    let senderEmail = '';
    try {
      if (item.from) {
        sender = item.from.displayName || '';
        senderEmail = item.from.emailAddress || '';
      }
    } catch (e) {
      console.warn('Could not get sender');
    }

    // Get attachments info
    const attachments: OutlookAttachment[] = [];
    try {
      if (item.attachments) {
        const atts = item.attachments;
        for (let i = 0; i < atts.length; i++) {
          attachments.push({
            id: atts[i].id || `att-${i}`,
            name: atts[i].name || `Attachment ${i+1}`,
            isFile: true,
            size: atts[i].size,
            contentType: atts[i].contentType,
          });
        }
      }
    } catch (e) {
      console.warn('Could not get attachments');
    }

    return {
      subject: subject || 'No Subject',
      body: '', // Will be fetched async
      sender: sender || 'Unknown Sender',
      senderEmail: senderEmail || '',
      attachments,
      receivedDate: new Date(),
    };
  } catch (error) {
    console.error('Error getting Outlook item data:', error);
    return null;
  }
};

// Async version to get full email data with body
export const getOutlookItemDataAsync = async (): Promise<OutlookItemData> => {
  return new Promise((resolve, reject) => {
    try {
      const item = Office.context.mailbox.item;
      
      if (!item) {
        reject(new Error('No Outlook item available'));
        return;
      }

      const result: OutlookItemData = {
        subject: '',
        body: '',
        sender: '',
        senderEmail: '',
        attachments: [],
      };

      // Get subject
      try {
        result.subject = item.subject || 'No Subject';
      } catch (e) {
        result.subject = 'No Subject';
      }

      // Get sender
      try {
        if (item.from) {
          result.sender = item.from.displayName || '';
          result.senderEmail = item.from.emailAddress || '';
        }
      } catch (e) {
        console.warn('Could not get sender');
      }

      // Get CC and TO
      try {
        if (item.cc) {
          result.cc = item.cc.map((recipient: any) => 
            recipient.displayName || recipient.emailAddress || ''
          ).filter(Boolean);
        }
        if (item.to) {
          result.to = item.to.map((recipient: any) => 
            recipient.displayName || recipient.emailAddress || ''
          ).filter(Boolean);
        }
      } catch (e) {
        console.warn('Could not get recipients');
      }

      // Get conversation ID
      try {
        if (item.conversationId) {
          result.conversationId = item.conversationId;
        }
      } catch (e) {
        console.warn('Could not get conversation ID');
      }

      // Get attachments
      try {
        if (item.attachments) {
          const atts = item.attachments;
          for (let i = 0; i < atts.length; i++) {
            result.attachments.push({
              id: atts[i].id || `att-${i}`,
              name: atts[i].name || `Attachment ${i+1}`,
              isFile: true,
              size: atts[i].size,
              contentType: atts[i].contentType,
            });
          }
        }
      } catch (e) {
        console.warn('Could not get attachments');
      }

      // Get body (async)
      if (item.body) {
        item.body.getAsync(Office.CoercionType.Text, (bodyResult) => {
          if (bodyResult.status === Office.AsyncResultStatus.Succeeded) {
            result.body = bodyResult.value || '';
            
            // Also get HTML body if needed
            item.body.getAsync(Office.CoercionType.Html, (htmlResult) => {
              if (htmlResult.status === Office.AsyncResultStatus.Succeeded) {
                result.bodyHtml = htmlResult.value || '';
              }
              resolve(result);
            });
          } else {
            resolve(result);
          }
        });
      } else {
        resolve(result);
      }
    } catch (error) {
      reject(error);
    }
  });
};

// Get attachments as base64 (for uploading to ProWorkflow)
export const getAttachmentContent = async (attachmentId: string): Promise<string | null> => {
  return new Promise((resolve) => {
    try {
      const item = Office.context.mailbox.item;
      if (!item) {
        resolve(null);
        return;
      }

      // Get attachment content
      if (item.getAttachmentContentAsync) {
        item.getAttachmentContentAsync(attachmentId, (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            resolve(result.value);
          } else {
            resolve(null);
          }
        });
      } else {
        resolve(null);
      }
    } catch (error) {
      console.error('Error getting attachment:', error);
      resolve(null);
    }
  });
};

// Get all attachments content
export const getAllAttachmentsContent = async (): Promise<Map<string, string>> => {
  const attachmentsMap = new Map<string, string>();
  
  try {
    const item = Office.context.mailbox.item;
    if (!item || !item.attachments) {
      return attachmentsMap;
    }

    for (const attachment of item.attachments) {
      const content = await getAttachmentContent(attachment.id);
      if (content) {
        attachmentsMap.set(attachment.id, content);
      }
    }
  } catch (error) {
    console.error('Error getting attachments content:', error);
  }

  return attachmentsMap;
};

// Clean email body (remove email signatures, replies, etc.)
export const cleanEmailBody = (body: string): string => {
  if (!body) return '';
  
  let cleaned = body;
  
  // Remove email signatures (common patterns)
  const signaturePatterns = [
    /--\s*$/m,
    /__\s*$/m,
    /Sent from my.*/i,
    /Get Outlook for.*/i,
    /Sent from.*/i,
    /________________________________/,
    /Disclaimer:.*/i,
    /Confidentiality Notice:.*/i,
  ];
  
  for (const pattern of signaturePatterns) {
    cleaned = cleaned.replace(pattern, '');
  }
  
  // Remove reply headers
  const replyPatterns = [
    /On.*wrote:/i,
    /-----Original Message-----/i,
    /From:.*/i,
    /Sent:.*/i,
    /To:.*/i,
    /Subject:.*/i,
  ];
  
  for (const pattern of replyPatterns) {
    cleaned = cleaned.replace(pattern, '');
  }
  
  // Trim extra whitespace
  cleaned = cleaned.trim();
  
  return cleaned;
};

// Generate task description from email
export const generateTaskDescription = (emailData: OutlookItemData): string => {
  let description = '';
  
  if (emailData.body) {
    description = cleanEmailBody(emailData.body);
  }
  
  // Add metadata if available
  const metadata = [];
  if (emailData.sender) {
    metadata.push(`From: ${emailData.sender} (${emailData.senderEmail})`);
  }
  if (emailData.receivedDate) {
    metadata.push(`Received: ${emailData.receivedDate.toLocaleString()}`);
  }
  if (emailData.subject && emailData.subject !== 'No Subject') {
    metadata.push(`Original Subject: ${emailData.subject}`);
  }
  if (emailData.cc && emailData.cc.length > 0) {
    metadata.push(`CC: ${emailData.cc.join(', ')}`);
  }
  
  if (metadata.length > 0) {
    description = `--- Email Metadata ---\n${metadata.join('\n')}\n\n--- Email Content ---\n${description}`;
  }
  
  return description;
};
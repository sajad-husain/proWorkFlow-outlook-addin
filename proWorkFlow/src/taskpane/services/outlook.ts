// Outlook Service - Extract data from current email

export interface OutlookItemData {
  subject: string;
  body: string;
  sender: string;
  senderEmail: string;
  attachments: OutlookAttachment[];
  receivedDate?: Date;
}

export interface OutlookAttachment {
  id: string;
  name: string;
  size?: number;
  contentType?: string;
  isFile?: boolean;
}

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
      subject = item.subject || '';
    } catch (e) {
      console.warn('Could not get subject');
    }

    // Get body (as text)
    let body = '';
    try {
      // For simplicity, get body as text
      if (item.body) {
        // You might want to use getAsync for HTML content
        body = item.body.getAsync('text', (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            return result.value;
          }
          return '';
        });
        // For synchronous access, we'll use item.body directly if available
        // In real implementation, use async methods
      }
    } catch (e) {
      console.warn('Could not get body');
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
        // In Outlook, attachments are available but accessing them
        // requires async methods. We'll store basic info.
        const atts = item.attachments;
        for (let i = 0; i < atts.length; i++) {
          attachments.push({
            id: atts[i].id || `att-${i}`,
            name: atts[i].name || `Attachment ${i+1}`,
            isFile: true,
          });
        }
      }
    } catch (e) {
      console.warn('Could not get attachments');
    }

    return {
      subject: subject || 'No Subject',
      body: body || '',
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

// Async version to get full email data
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
      result.subject = item.subject || 'No Subject';

      // Get sender
      if (item.from) {
        result.sender = item.from.displayName || '';
        result.senderEmail = item.from.emailAddress || '';
      }

      // Get body (async)
      if (item.body) {
        item.body.getAsync(Office.CoercionType.Text, (bodyResult) => {
          if (bodyResult.status === Office.AsyncResultStatus.Succeeded) {
            result.body = bodyResult.value || '';
          }
          resolve(result);
        });
      } else {
        resolve(result);
      }
    } catch (error) {
      reject(error);
    }
  });
};
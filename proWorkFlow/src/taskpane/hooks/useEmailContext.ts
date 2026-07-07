import { useState, useEffect } from 'react';

export interface EmailData {
  subject: string;
  body: string;
  from: string;
  to?: string;
  attachments?: any[];
  itemId?: string;
}

export const useEmailContext = () => {
  const [emailData, setEmailData] = useState<EmailData | null>(null);
  const [loading, setLoading] = useState<boolean>(true);

  useEffect(() => {
    const loadEmailData = async () => {
      try {
        const item = Office.context.mailbox.item;
        
        // Get subject (sync)
        const subject = item.subject || '';
        
        // Get from
        const from = item.from ? item.from.emailAddress || '' : '';
        
        // Get body (async)
        let body = '';
        if (item.body) {
          const bodyResult = await new Promise<Office.AsyncResult<string>>((resolve) => {
            item.body?.getAsync('text', { asyncContext: null }, (result) => {
              resolve(result);
            });
          });
          body = bodyResult.value || '';
        }
        
        // Get attachments
        const attachments = item.attachments || [];
        
        setEmailData({
          subject,
          body,
          from,
          attachments,
          itemId: item.itemId,
        });
        setLoading(false);
      } catch (error) {
        console.error('Error loading email context:', error);
        setLoading(false);
      }
    };

    if (Office.context.mailbox) {
      loadEmailData();
    } else {
      // Fallback for testing outside Outlook
      setEmailData({
        subject: 'Re: Do you have a plug in for microsoft outlook...',
        body: 'Hey there Joshua,\n\nYou can find that from your client area, but here it is also:',
        from: 'support@proworkflow.com',
        attachments: [],
      });
      setLoading(false);
    }
  }, []);

  return { emailData, loading };
};
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("unpinButton").onclick = unpinAllEmails;
    document.getElementById("countButton").onclick = countPinnedEmails;
  }
});

function showStatus(message, type = 'info') {
  const statusDiv = document.getElementById('status');
  statusDiv.textContent = message;
  statusDiv.className = `status ${type}`;
  statusDiv.style.display = 'block';
}

function disableButtons(disabled) {
  document.getElementById('unpinButton').disabled = disabled;
  document.getElementById('countButton').disabled = disabled;
}

async function countPinnedEmails() {
  showStatus('Counting pinned emails...', 'info');
  disableButtons(true);
  
  try {
    // Use Graph API approach first, fallback to EWS if needed
    if (Office.context.mailbox.restUrl) {
      await countWithGraphAPI();
    } else {
      await countWithEWS();
    }
  } catch (error) {
    console.error('Error counting emails:', error);
    showStatus('Failed to count pinned emails: ' + error.message, 'error');
  } finally {
    disableButtons(false);
  }
}

async function countWithGraphAPI() {
  try {
    Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, (result) => {
      if (result.status === 'succeeded') {
        const accessToken = result.value;
        
        fetch(`${Office.context.mailbox.restUrl}/v2.0/me/mailFolders/inbox/messages?$filter=flag/flagStatus eq 'flagged'&$count=true&$top=1`, {
          method: 'GET',
          headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Accept': 'application/json'
          }
        })
        .then(response => response.json())
        .then(data => {
          const count = data['@odata.count'] || 0;
          showStatus(`Found ${count} pinned emails`, 'info');
        })
        .catch(error => {
          console.error('Graph API error:', error);
          countWithEWS();
        });
      } else {
        countWithEWS();
      }
    });
  } catch (error) {
    countWithEWS();
  }
}

async function countWithEWS() {
  const ewsQuery = `<?xml version="1.0" encoding="utf-8"?>
    <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                   xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
                   xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
                   xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
      <soap:Header>
        <t:RequestServerVersion Version="Exchange2013" />
      </soap:Header>
      <soap:Body>
        <m:FindItem Traversal="Shallow">
          <m:ItemShape>
            <t:BaseShape>IdOnly</t:BaseShape>
          </m:ItemShape>
          <m:Restriction>
            <t:IsEqualTo>
              <t:FieldURI FieldURI="message:Flag" />
              <t:FieldURIOrConstant>
                <t:Constant Value="2" />
              </t:FieldURIOrConstant>
            </t:IsEqualTo>
          </m:Restriction>
          <m:ParentFolderId>
            <t:DistinguishedFolderId Id="inbox" />
          </m:ParentFolderId>
        </m:FindItem>
      </soap:Body>
    </soap:Envelope>`;

  Office.context.mailbox.makeEwsRequestAsync(ewsQuery, (result) => {
    if (result.status === 'succeeded') {
      const parser = new DOMParser();
      const xmlDoc = parser.parseFromString(result.value, "text/xml");
      const items = xmlDoc.getElementsByTagName("t:ItemId");
      showStatus(`Found ${items.length} pinned emails using EWS`, 'info');
    } else {
      showStatus('Failed to count emails: ' + result.error.message, 'error');
    }
  });
}

async function unpinAllEmails() {
  showStatus('Starting unpin process...', 'info');
  disableButtons(true);
  
  try {
    // Use Graph API approach first, fallback to EWS if needed
    if (Office.context.mailbox.restUrl) {
      await unpinWithGraphAPI();
    } else {
      await unpinWithEWS();
    }
  } catch (error) {
    console.error('Error unpinning emails:', error);
    showStatus('Failed to unpin emails: ' + error.message, 'error');
  } finally {
    disableButtons(false);
  }
}

async function unpinWithGraphAPI() {
  try {
    Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, async (result) => {
      if (result.status === 'succeeded') {
        const accessToken = result.value;
        let processedCount = 0;
        let hasMore = true;
        let skipToken = '';
        
        while (hasMore) {
          const url = `${Office.context.mailbox.restUrl}/v2.0/me/mailFolders/inbox/messages?$filter=flag/flagStatus eq 'flagged'&$top=100${skipToken ? '&$skiptoken=' + skipToken : ''}`;
          
          const response = await fetch(url, {
            method: 'GET',
            headers: {
              'Authorization': `Bearer ${accessToken}`,
              'Accept': 'application/json'
            }
          });
          
          const data = await response.json();
          
          if (data.value && data.value.length > 0) {
            // Process emails in batches
            for (const email of data.value) {
              try {
                await fetch(`${Office.context.mailbox.restUrl}/v2.0/me/messages/${email.id}`, {
                  method: 'PATCH',
                  headers: {
                    'Authorization': `Bearer ${accessToken}`,
                    'Content-Type': 'application/json'
                  },
                  body: JSON.stringify({
                    flag: {
                      flagStatus: 'notFlagged'
                    }
                  })
                });
                processedCount++;
                
                if (processedCount % 10 === 0) {
                  showStatus(`Unpinned ${processedCount} emails...`, 'info');
                }
              } catch (emailError) {
                console.error('Error unpinning individual email:', emailError);
              }
            }
            
            // Check for more results
            hasMore = !!data['@odata.nextLink'];
            if (hasMore && data['@odata.nextLink'].includes('$skiptoken=')) {
              skipToken = data['@odata.nextLink'].split('$skiptoken=')[1];
            }
          } else {
            hasMore = false;
          }
        }
        
        showStatus(`Successfully unpinned ${processedCount} emails!`, 'success');
      } else {
        await unpinWithEWS();
      }
    });
  } catch (error) {
    console.error('Graph API error:', error);
    await unpinWithEWS();
  }
}

async function unpinWithEWS() {
  let processedCount = 0;
  let offset = 0;
  const batchSize = 100;
  let hasMore = true;

  while (hasMore) {
    const findQuery = `<?xml version="1.0" encoding="utf-8"?>
      <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                     xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
                     xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
                     xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
        <soap:Header>
          <t:RequestServerVersion Version="Exchange2013" />
        </soap:Header>
        <soap:Body>
          <m:FindItem Traversal="Shallow">
            <m:ItemShape>
              <t:BaseShape>IdOnly</t:BaseShape>
            </m:ItemShape>
            <m:IndexedPageItemView MaxEntriesReturned="${batchSize}" Offset="${offset}" BasePoint="Beginning" />
            <m:Restriction>
              <t:IsEqualTo>
                <t:FieldURI FieldURI="message:Flag" />
                <t:FieldURIOrConstant>
                  <t:Constant Value="2" />
                </t:FieldURIOrConstant>
              </t:IsEqualTo>
            </m:Restriction>
            <m:ParentFolderId>
              <t:DistinguishedFolderId Id="inbox" />
            </m:ParentFolderId>
          </m:FindItem>
        </soap:Body>
      </soap:Envelope>`;

    await new Promise((resolve, reject) => {
      Office.context.mailbox.makeEwsRequestAsync(findQuery, (findResult) => {
        if (findResult.status === 'succeeded') {
          const parser = new DOMParser();
          const xmlDoc = parser.parseFromString(findResult.value, "text/xml");
          const items = xmlDoc.getElementsByTagName("t:ItemId");
          
          if (items.length === 0) {
            hasMore = false;
            resolve();
            return;
          }

          // Create update query for found items
          let updateQuery = `<?xml version="1.0" encoding="utf-8"?>
            <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                           xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
                           xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
                           xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
              <soap:Header>
                <t:RequestServerVersion Version="Exchange2013" />
              </soap:Header>
              <soap:Body>
                <m:UpdateItem ConflictResolution="AutoResolve" MessageDisposition="SaveOnly">
                  <m:ItemChanges>`;

          for (let i = 0; i < items.length; i++) {
            const itemId = items[i].getAttribute('Id');
            const changeKey = items[i].getAttribute('ChangeKey');
            updateQuery += `
                    <t:ItemChange>
                      <t:ItemId Id="${itemId}" ChangeKey="${changeKey}" />
                      <t:Updates>
                        <t:SetItemField>
                          <t:FieldURI FieldURI="message:Flag" />
                          <t:Message>
                            <t:Flag>
                              <t:FlagStatus>NotFlagged</t:FlagStatus>
                            </t:Flag>
                          </t:Message>
                        </t:SetItemField>
                      </t:Updates>
                    </t:ItemChange>`;
          }

          updateQuery += `
                  </m:ItemChanges>
                </m:UpdateItem>
              </soap:Body>
            </soap:Envelope>`;

          Office.context.mailbox.makeEwsRequestAsync(updateQuery, (updateResult) => {
            if (updateResult.status === 'succeeded') {
              processedCount += items.length;
              showStatus(`Unpinned ${processedCount} emails...`, 'info');
              offset += batchSize;
              resolve();
            } else {
              reject(new Error('Update failed: ' + updateResult.error.message));
            }
          });
        } else {
          reject(new Error('Find failed: ' + findResult.error.message));
        }
      });
    });
  }

  showStatus(`Successfully unpinned ${processedCount} emails using EWS!`, 'success');
}

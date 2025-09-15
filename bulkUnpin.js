Office.onReady(() => {
  document.getElementById("unpinButton").onclick = async () => {
    try {
      const response = await Office.context.mailbox.makeEwsRequestAsync(`
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
              <m:IndexedPageItemView MaxEntriesReturned="100" Offset="0" BasePoint="Beginning" />
              <m:ParentFolderId>
                <t:DistinguishedFolderId Id="inbox" />
              </m:ParentFolderId>
            </m:FindItem>
          </soap:Body>
        </soap:Envelope>
      `);
      console.log("Response:", response);
      alert("Unpin operation initiated.");
    } catch (error) {
      console.error("Error:", error);
      alert("Failed to unpin emails.");
    }
  };
});

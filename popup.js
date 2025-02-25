document.getElementById('openEditable').addEventListener('click', async () => {
    const statusDiv = document.getElementById('status');
    try {
        // Get the current active tab
        const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
        
        if (!tab || !tab.url.includes('sharepoint.com')) {
            statusDiv.textContent = 'Not a SharePoint document';
            return;
        }

        const url = new URL(tab.url);
        
        // If it's already in edit mode, don't do anything
        if (url.searchParams.get('action') === 'edit') {
            statusDiv.textContent = 'Document is already in edit mode';
            return;
        }

        // Check if this is the non-editable view (contains Doc.aspx)
        if (url.pathname.includes('/_layouts/15/Doc.aspx')) {
            const sourceDoc = url.searchParams.get('sourcedoc');
            const file = url.searchParams.get('file');
            
            if (!sourceDoc || !file) {
                statusDiv.textContent = 'Unable to get document information';
                return;
            }

            // Simpler URL transformation
            const newUrl = new URL(tab.url);
            newUrl.searchParams.set('action', 'edit');
            newUrl.searchParams.delete('mobileredirect');
            newUrl.searchParams.delete('wdLOR');

            statusDiv.textContent = 'Opening document in edit mode...';
            await chrome.tabs.update(tab.id, { url: newUrl.toString() });
        } else {
            // For other SharePoint URLs, just try to add edit parameter
            const newUrl = new URL(tab.url);
            newUrl.searchParams.set('action', 'edit');
            
            statusDiv.textContent = 'Opening document in edit mode...';
            await chrome.tabs.update(tab.id, { url: newUrl.toString() });
        }
    } catch (error) {
        statusDiv.textContent = 'Error: ' + error.message;
    }
}); 
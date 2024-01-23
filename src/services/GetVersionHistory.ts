const siteUrl = "https://xvzms.sharepoint.com/";
const libraryName = "Documents";
const fileName = "TestDocument.docx";

// Get the current item (document) ID
const getItemUrl = `${siteUrl}/_api/web/lists/getbytitle('${libraryName}')/items?$filter=FileLeafRef eq '${fileName}'&$select=ID`;

fetch(getItemUrl, {
    method: 'GET',
    headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-Type': 'application/json',
    },
})
.then(response => response.json())
.then(data => {
    if (data.value.length > 0) {
        const itemId = data.value[0].ID;
        
        // Get the version history of the item
        const getVersionHistoryUrl = `${siteUrl}/_api/web/lists/getbytitle('${libraryName}')/items(${itemId})/versions`;

        fetch(getVersionHistoryUrl, {
            method: 'GET',
            headers: {
                'Accept': 'application/json;odata=nometadata',
                'Content-Type': 'application/json',
            },
        })
        .then(response => response.json())
        .then(versionData => {
            // Assume you want to download the first version, change index as needed
            const versionIndex = 0;
            const versionId = versionData.value[versionIndex].VersionId;

            // Download the specific version of the document
            const downloadUrl = `${siteUrl}/_layouts/15/download.aspx?SourceUrl=${encodeURIComponent(`/${libraryName}/${fileName}`)}&FldEdit=0&ver=${versionId}`;

            // Open the download link in a new window or redirect the user to trigger the download
            window.open(downloadUrl, '_blank');
        })
        .catch(error => {
            console.error('Error fetching version history:', error);
        });
    } else {
        console.error('Document not found');
    }
})
.catch(error => {
    console.error('Error fetching item ID:', error);
});
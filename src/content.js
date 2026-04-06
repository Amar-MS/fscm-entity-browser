// content.js — Runs inside F&SCM pages; proxies OData API calls for the side panel.
// Because it runs in the page context, it inherits the user's existing auth session.

chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
  if (message.type !== 'FSCM_API_REQUEST') return false;

  const { url, method = 'GET', body } = message;

  const fetchOptions = {
    method,
    headers: {
      'Accept': 'application/json;odata.metadata=minimal',
      'OData-MaxVersion': '4.0',
      'OData-Version': '4.0'
    },
    credentials: 'same-origin'
  };

  if (body) {
    fetchOptions.headers['Content-Type'] = 'application/json';
    fetchOptions.body = JSON.stringify(body);
  }

  // For $metadata, we accept XML
  if (url.includes('$metadata')) {
    fetchOptions.headers['Accept'] = 'application/xml';
  }

  fetch(url, fetchOptions)
    .then(async (response) => {
      if (!response.ok) {
        const text = await response.text();
        throw new Error(`HTTP ${response.status}: ${text.substring(0, 300)}`);
      }

      const contentType = response.headers.get('content-type') || '';
      if (contentType.includes('xml')) {
        const text = await response.text();
        return { type: 'xml', data: text };
      }

      const json = await response.json();
      return { type: 'json', data: json };
    })
    .then((result) => sendResponse({ success: true, ...result }))
    .catch((err) => sendResponse({ success: false, error: err.message }));

  return true; // Keep channel open for async response
});

// background.js — Open full browser tab on toolbar icon click

chrome.action.onClicked.addListener((tab) => {
  if (tab.id && tab.url?.includes('.operations.dynamics.com')) {
    chrome.storage.session.set({
      activeFscmTabId:  tab.id,
      activeFscmTabUrl: tab.url
    });
  }
  chrome.tabs.create({ url: chrome.runtime.getURL('src/sidepanel.html') });
});

// Track active F&SCM tab so the browser page always knows which env to call
chrome.tabs.onActivated.addListener(async ({ tabId }) => {
  try {
    const tab = await chrome.tabs.get(tabId);
    if (tab.url?.includes('.operations.dynamics.com')) {
      chrome.storage.session.set({
        activeFscmTabId:  tabId,
        activeFscmTabUrl: tab.url
      });
    }
  } catch (_) {}
});

chrome.tabs.onUpdated.addListener((tabId, changeInfo, tab) => {
  if (changeInfo.status === 'complete' && tab.url?.includes('.operations.dynamics.com')) {
    chrome.storage.session.set({
      activeFscmTabId:  tabId,
      activeFscmTabUrl: tab.url
    });
  }
});

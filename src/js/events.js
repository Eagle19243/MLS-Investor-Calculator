chrome.browserAction.onClicked.addListener(onBrowserActionClicked);
chrome.tabs.onActivated.addListener(onTabActivated);
chrome.tabs.onUpdated.addListener(onTabUpdated);

function onTabActivated(activeInfo) {
    chrome.tabs.get(activeInfo.tabId, (tab) => {
        determineExtensionAvailability(tab.url);
    });
}

function onTabUpdated(tabId, changeInfo, tab) {
    determineExtensionAvailability(tab.url);
}

function onBrowserActionClicked(tab) {
    chrome.tabs.query({active: true, currentWindow: true}, (tabs) => {
        chrome.tabs.sendMessage(tabs[0].id, {action: 'Show_Popup'}, (response) => {});
    });
}

function determineExtensionAvailability(url) {
    if (!url || (!SPINURLRegex.test(url) && !MATRIXURLRegex.test(url))) {
        chrome.browserAction.setIcon({path: 'images/icon-gray-48.png'});
        chrome.browserAction.disable();
    } else {
        chrome.browserAction.setIcon({path: 'images/icon-48.png'});
        chrome.browserAction.enable();
    }
}


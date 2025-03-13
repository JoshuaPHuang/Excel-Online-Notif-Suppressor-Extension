chrome.storage.local.get(["xlpbLogs"], function(result) {
    const logOutput = document.getElementById('log-output');
    logOutput.textContent = result.cimsLogs.join("\n") || 'No logs available.';
});
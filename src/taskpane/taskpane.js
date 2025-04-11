import { saveAs } from 'file-saver';

Office.onReady((info) => {
  // Check if we're in Outlook
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("save-button").onclick = saveEmailAsJson;
  }
});

/**
 * Gets the current email item and extracts its content to save as JSON
 */
function saveEmailAsJson() {
  const statusElement = document.getElementById("status");
  statusElement.innerText = "Processing...";

  try {
    // Get the current item (email)
    const item = Office.context.mailbox.item;
    
    if (!item) {
      statusElement.innerText = "No email selected";
      return;
    }

    // Create an object to store email data
    const emailData = {
      subject: item.subject,
      sender: item.sender ? item.sender.emailAddress : "Unknown",
      receivedTime: item.dateTimeCreated ? new Date(item.dateTimeCreated).toISOString() : null,
      bodyContent: null
    };

    // Get the email body content
    item.body.getAsync(Office.CoercionType.Text, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        // Add the body text to our data object
        emailData.bodyContent = result.value;
        
        // Convert the email data to a JSON string
        const jsonData = JSON.stringify(emailData, null, 2);
        
        // Create a Blob from the JSON string
        const blob = new Blob([jsonData], { type: "application/json" });
        
        // Generate a filename using the subject (or a default name if no subject)
        const filename = `${emailData.subject || "email"}_${new Date().getTime()}.json`;
        
        // Save the file
        saveAs(blob, filename);
        
        statusElement.innerText = "Email saved as JSON successfully!";
      } else {
        statusElement.innerText = `Error getting email body: ${result.error.message}`;
      }
    });
  } catch (error) {
    statusElement.innerText = `Error: ${error.message}`;
  }
}

// Fallback function if Office.js is not available
function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}

const g = getGlobal(); 
/**
 * Logs a message to the console with a [DEBUG] prefix.
 * @param {string} message - The message to log.
 */
function log(message) {
  console.log(`[DEBUG] ${message}`);
}

// A private variable to hold our configuration data.
// A Map is used for efficient key-value pair storage and retrieval.
let configMap = new Map();

/**
 * Fetches the config.json file from the specified URL,
 * parses the JSON, and loads the key-value pairs into a map.
 * This is an asynchronous operation.
 * @param {string} url The URL of the config.json file.
 * @returns {Promise<void>} A promise that resolves when the config is loaded.
 */
async function loadConfig(url) {
  try {
    const response = await fetch(url);
    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }
    const parsedData = await response.json();

    // Unpack the key-value pairs into the global map
    for (const key in parsedData) {
      configMap.set(key, parsedData[key]);
    }
  } catch (e) {
    console.error(`Failed to load configuration: ${e.message}`);
    throw e;
  }
}

/**
 * Retrieves a value from the configuration map using its key.
 * @param {string} key The key to look up in the config.
 * @returns {*} The value associated with the key, or undefined if the key is not found.
 */
function getVar(key) {
  if (configMap.size !== 0) {
    const value = configMap.get(key);
    // If the value is an array, join its elements with a newline character.
    if (Array.isArray(value)) {
      return value.join(String.fromCharCode(10));
    }
    return value;
  }
  return "";
}

/**
 * Removes the leading and trailing triple-backtick 'html' fences from a string.
 * This is useful for cleaning up code blocks formatted with Markdown.
 *
 * @param {string} content - The string that may contain the HTML fences.
 * @returns {string} The string with the fences removed.
 */
function removeHtmlFences(content) {
  const fence = '```html' + String.fromCharCode(10);
  
  if (content.startsWith(fence) && content.endsWith('```')) {
    return content.slice(fence.length, -3);
  }
  
  return content;
}

/**
 * Sets the metadata for the email pane.
 * @param {string} sentiment - The sentiment of the email.
 * @param {string} urgency - The urgency of the email.
 * @param {string} intention - The intention of the email.
 * @returns {void}
 */
function setMetaData(sentiment, urgency, intention) {
  // Logs the retrieved metadata
  log(`Email Metadata - Sentiment: ${sentiment}, Urgency: ${urgency}, Intention: ${intention}`);

  // Get the indicator element
  const urgencyIndicator = document.getElementById("urgencyIndicator");

  // Updates the UI with the retrieved metadata
  document.getElementById("sentiment").textContent = sentiment;
  document.getElementById("urgency").textContent = urgency;
  document.getElementById("intention").textContent = intention;

  // Update urgency indicator color
  if (urgencyIndicator) {
    urgencyIndicator.classList.remove("urgency-high", "urgency-medium", "urgency-low"); // Reset state
    const processedUrgency = urgency.trim().toLowerCase();

    if (processedUrgency.includes("high")) {
      urgencyIndicator.classList.add("urgency-high");
    } else if (processedUrgency.includes("medium")) {
      urgencyIndicator.classList.add("urgency-medium");
    } else if (processedUrgency.includes("low")) {
      urgencyIndicator.classList.add("urgency-low");
    }
  }
  return;
}

/**
 * Updates the document body styles based on the Office theme.
 * @param {Object} theme - The Office theme object.
 */
function updateThemeStyles(theme) {
  if (theme) {
    document.documentElement.style.setProperty("--bg-color", theme.bodyBackgroundColor);
    document.documentElement.style.setProperty("--text-color", theme.bodyForegroundColor);
  }
}

/**
 * Registers the Office theme change event handler.
 */
function registerThemeChangeHandler() {
  const mailbox = Office.context?.mailbox;
  if (mailbox) {
    mailbox.addHandlerAsync(
      Office.EventType.OfficeThemeChanged,
      (eventArgs) => {
        const theme = eventArgs.officeTheme;
        updateThemeStyles(theme);
      },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          console.error("Failed to register theme change handler:", result.error.message);
        } else {
          log("Theme change handler registered successfully.");
        }
      }
    );
  }
}

/**
 * Calls the custom API to generate a response.
 * Includes a retry mechanism with exponential backoff for rate limit errors.
 * @param {string} userQuery - The user's query to send to the API.
 * @returns {Promise<string>} A promise that resolves with the generated text.
 */
async function callCustomEndpoint(userQuery) {
  const apiUrl = getVar("customendpoint_url");
  const payload = {
    query: userQuery,
  };

  let attempts = 0;
  const maxAttempts = 5;
  const baseDelay = 1000;

  while (attempts < maxAttempts) {
    try {
      log(`Attempting to call custom API. Attempt ${attempts + 1} of ${maxAttempts}.`);
      const response = await fetch(apiUrl, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(payload),
      });
      log(`Attempt ${attempts + 1}: API response status is ${response.status}`);

      // Handles rate limiting with exponential backoff
      if (response.status === 429) {
        const delay = baseDelay * Math.pow(2, attempts) + Math.random() * 1000;
        attempts++;
        log(`Rate limit exceeded. Retrying in ${delay}ms.`);
        await new Promise((res) => setTimeout(res, delay));
        continue;
      }

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`API call failed with status ${response.status}: ${errorText}`);
      }

      const result = await response.json();
      // Checks for a valid response format and returns the generated text
      if (result) {
        log("API call successful.");
        return result;
      } else {
        throw new Error("Invalid response format from API.");
      }
    } catch (error) {
      attempts++;
      log(`Attempt ${attempts}: an error occurred. ${error.message}`);
      if (attempts >= maxAttempts) {
        throw new Error(`Failed to call custom API after ${maxAttempts} attempts.`);
      }
    }
  }
}

/**
 * Calls the Gemini API to generate a response.
 * Includes a retry mechanism with exponential backoff for rate limit errors.
 * @param {string} userQuery - The user's query to send to the API.
 * @param {string} system_instruction - The system instruction for the API.
 * @returns {Promise<string>} A promise that resolves with the generated text.
 */
async function callGeminiAPI(userQuery, system_instruction) {
  const apiKey = getVar("google_api_key"); // Replace with your actual API key or use a secure method to store it
  if (!apiKey) {
    throw new Error("GEMINI_API_KEY is not set in environment variables.");
  }
  const apiUrl = `${getVar("endpoint_url")}${apiKey}`;
  const payload = {
    contents: [
      {
        parts: [
          {
            text: userQuery,
          },
        ],
      },
    ],
    tools: [
      {
        google_search: {},
      },
    ],
    systemInstruction: {
      parts: [
        {
          text: system_instruction,
        },
      ],
    },
  };

  let attempts = 0;
  const maxAttempts = 5;
  const baseDelay = 1000;

  while (attempts < maxAttempts) {
    try {
      const response = await fetch(apiUrl, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(payload),
      });
      log(`Attempt ${attempts + 1}: API response status is ${response.status}`);

      // Handles rate limiting with exponential backoff
      if (response.status === 429) {
        const delay = baseDelay * Math.pow(2, attempts) + Math.random() * 1000;
        attempts++;
        log(`Rate limit exceeded. Retrying in ${delay}ms.`);
        await new Promise((res) => setTimeout(res, delay));
        continue;
      }

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`API call failed with status ${response.status}: ${errorText}`);
      }

      const result = await response.json();
      const candidate = result.candidates?.[0];

      // Checks for a valid response format and returns the generated text
      if (candidate && candidate.content?.parts?.[0]?.text) {
        log("API call successful.");
        return candidate.content.parts[0].text;
      } else {
        throw new Error("Invalid response format from API.");
      }
    } catch (error) {
      attempts++;
      log(`Attempt ${attempts}: an error occurred. ${error.message}`);
      if (attempts >= maxAttempts) {
        throw new Error(`Failed to call Gemini API after ${maxAttempts} attempts.`);
      }
    }
  }
}

async function getEmailBody(item) {
  return new Promise((resolve, reject) => {
    if (item.body?.getAsync) {
      item.body.getAsync(Office.CoercionType.Text, (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
          resolve(asyncResult.value);
        } else {
          reject(asyncResult.error);
        }
      });
    } else {
      resolve(""); // Resolve with an empty string if no body exists.
    }
  });
}

/**
 * The main entry point for the Office Add-in.
 * This function runs when the Office document is ready.
 */
Office.onReady(async (info) => {
  log("Office.js is ready.");
  const configUrl = "config/config.json";
  
  // Load the config file from the specified URL
  try {
    await loadConfig(configUrl);
    log("Configuration loaded successfully.");
  } catch (error) {
    console.error(`Failed to load configuration: ${error.message}`);
    log(`Failed to load configuration: ${error.message}`);
  }

  if (info) {
    log(`Add-in is running in ${info.host} on ${info.platform}.`);
  }

  // Set initial theme
  const theme = Office.context.officeTheme;
  if (theme) {
    updateThemeStyles(theme);
  }

  // Register theme change handler after onReady
  registerThemeChangeHandler();

  const item = Office.context?.mailbox?.item;

  const helpdeskPrompt = getVar("helpdeskPrompt");
  const sentimentPrompt = getVar("sentimentPrompt");
  const urgencyPrompt = getVar("urgencyPrompt");
  const intentionPrompt = getVar("intentionPrompt");
  const customEndpointUrl = getVar("customendpoint_url");

  if (!item) {
    log("No mail item found. Add-in may be running in an unsupported context.");
    // Hides the main content and disables buttons
    document.getElementById("sentimentContent").classList.add("hidden");
    document.getElementById("btnQuickReply").disabled = true;
    // Displays a message to the user
    document.getElementById("responseContainer").textContent =
      "This add-in only works with email messages.";
    return; // Stop execution if no item is available
  }

  log(`Item type: ${item.itemType}`);

  // Checks the mode (Read or Compose) of the mail item
  const isReadMode = item.itemType === Office.MailboxEnums.ItemType.Message;
  const isComposeMode = item.itemType === Office.MailboxEnums.ItemType.MessageCompose;

  // Declaring draft variable here to be available to all listeners
  let draft = "";

  if (isReadMode) {
    log("In READ mode.");
    // Shows the read mode UI and enables the buttons
    document.getElementById("sentimentContent").classList.remove("hidden");
    document.getElementById("btnQuickReply").disabled = false;

    // Wait for the email body to be retrieved first
    const emailBody = (await getEmailBody(item)).trim();

    // Now that the 'preview' element is populated, you can safely proceed.
    document.getElementById("responseContainer").innerHTML =
      "Analyzing email content, please wait...";

    let name = "";
    if (item.from) {
      name = item.from.displayName || item.from.emailAddress;
    }

    try {
      if (customEndpointUrl !== "") {
        // Retrieves additional email metadata asynchronously
        const query = {
          fromEmailAddress: name,
          subject: item.subject ? item.subject : "Unknown",
          body: emailBody, // Use the variable with the body content
        };
        const responseData = await callCustomEndpoint(query);
        log(`Call response: ${responseData}`);

        let sentiment = responseData.response.metadata.email_sentiment;
        let urgency = responseData.response.metadata.email_urgency;
        let intention = responseData.response.metadata.email_intention;
        draft = removeHtmlFences(responseData.response.answer.email_draft);
        /*         
        let htmlResponse = "Analysis complete. I am ready to help you generate a draft.";
        let comments = responseData.response.metadata.email_review_comments;
        if (comments && comments !== "No further comments.") {
          htmlResponse += // Maybe use this one day...
            "Analysis complete. I am ready to help you generate a draft."+
            "<br><br>By the way, some review comments were identified that might help you."+
            "<br><br><b>Comments:</b><br>" + comments;
        }  
        */
        document.getElementById("responseContainer").innerHTML =
          "Analysis complete. Please click 'Quick Reply' generate a draft.";
        setMetaData(sentiment, urgency, intention);
      } else {
        // Retrieves additional email metadata asynchronously
        let sentiment = "Unknown";
        let urgency = "Unknown";
        let intention = "Unknown";

        // Calls the Gemini API to analyze sentiment
        try {
          const prompt = `From: ${name}\nBody: ${emailBody}`;
          sentiment = await callGeminiAPI(prompt, sentimentPrompt);
          log(`Sentiment analysis result: ${sentiment}`);
        } catch (error) {
          log(`Error analyzing sentiment: ${error.message}`);
        }
        // Calls the Gemini API to analyze urgency
        try {
          const prompt = `
            From: ${document.getElementById("from").textContent}
            Body: ${emailBody}`;
          urgency = await callGeminiAPI(prompt, urgencyPrompt);
          log(`Urgency analysis result: ${urgency}`);
        } catch (error) {
          console.error(`Error analyzing urgency: ${error.message}`);
          log(`Error analyzing urgency: ${error.message}`);
        }
        // Calls the Gemini API to analyze intention
        try {
          const prompt = `
            From: ${document.getElementById("from").textContent}
            Body: ${emailBody}`;
          intention = await callGeminiAPI(prompt, intentionPrompt);
          log(`Intention analysis result: ${intention}`);
        } catch (error) {
          console.error(`Error analyzing intention: ${error.message}`);
          log(`Error analyzing intention: ${error.message}`);
        }
        // ... (rest of Gemini API calls for urgency and intention)
        document.getElementById("responseContainer").innerHTML =
          "Analysis complete. Please click 'Quick Reply' generate a draft.";

        setMetaData(sentiment, urgency, intention);
      }
    } catch (error) {
      log(`Error calling endpoint: ${error.message}`);
      console.error(`Error calling endpoint: ${error.message}`);
    }

    // Event listener for the 'Quick Reply' button
    document.getElementById("btnQuickReply").addEventListener("click", async () => {
      // Disables the button and shows a loading message
      document.getElementById("btnQuickReply").disabled = true;
      document.getElementById("responseContainer").innerHTML = "Generating draft...";

      try {
        if (customEndpointUrl === "") {
          const prompt = `From: ${name}\nBody: ${emailBody}`;
          // Calls the Gemini API to generate the draft
          const generatedText = await callGeminiAPI(prompt, helpdeskPrompt);
          draft = generatedText;
        }
      } catch (error) {
        draft = "<html><body><p>Error generating draft. Please try again.</p></body></html>";
        log(`Error generating draft: ${error.message}`);
        console.error(`Error generating draft: ${error.message}`);
      } finally {
        // Re-enables the button
        document.getElementById("btnQuickReply").disabled = false;
        document.getElementById("responseContainer").innerHTML = "Draft generated successfully.";
      }
      // Displays the generated draft in a reply form with HTML coercion
      if (item.displayReplyAllForm) {
        item.displayReplyAllForm(draft, {
          coercionType: Office.CoercionType.Html,
        });
      }
    });
  } else if (isComposeMode) {
    log("In COMPOSE mode.");
    // Hides the main content and disables buttons for compose mode
    document.getElementById("sentimentContent").classList.add("hidden");
    document.getElementById("btnQuickReply").disabled = true;
    document.getElementById("responseContainer").innerHTML =
      "This functionality is not available in compose mode.";
  } else {
    log("In an unsupported mode.");
    // Hides the main content and disables buttons for unsupported modes
    document.getElementById("sentimentContent").classList.add("hidden");
    document.getElementById("btnQuickReply").disabled = true;
    document.getElementById("responseContainer").innerHTML =
      "This add-in only works with email messages.";
  }
});

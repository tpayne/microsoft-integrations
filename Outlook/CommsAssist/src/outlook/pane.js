// Apply Office theme (if available) and keep variables consistent with CSS names.
function applyOfficeThemeVars(theme) {
  if (!theme) return;

  // Hardcoded defaults that mirror your CSS :root light theme
  const FALLBACKS = {
    bg: "#ffffff",
    surface: "#fbfcfd",
    text: "#111827",
    border: "#e6e9ee",
    muted: "#6b7280",
    accent: "#0f64ff",
    shadow: "0 1px 2px rgba(16,24,40,0.04)"
  };

  // Helper: read a CSS variable safely with multiple fallbacks
  function readCssVar(name) {
    // 1) Prefer window.getComputedStyle if available
    try {
      if (typeof window !== "undefined" && typeof window.getComputedStyle === "function") {
        const comp = window.getComputedStyle(document.documentElement);
        const val = comp.getPropertyValue(name);
        if (val) return val.trim();
      }
    } catch {
      // ignore and try next fallback
    }

    // 2) Fallback to inline style set on documentElement
    try {
      const inline = document.documentElement.style.getPropertyValue(name);
      if (inline) return inline.trim();
    } catch {
      // ignore and try next fallback
    }

    // 3) Final fallback to hardcoded token (strip leading -- and return)
    const key = name.replace(/^--/, "");
    return FALLBACKS[key] || "";
  }

  // Read current CSS tokens safely
  const currentBg = readCssVar("--bg") || FALLBACKS.bg;
  const currentText = readCssVar("--text") || FALLBACKS.text;

  // Prefer explicit Office theme values when present, otherwise keep current tokens
  const bodyBg = (theme.bodyBackgroundColor && String(theme.bodyBackgroundColor).trim()) ? theme.bodyBackgroundColor : currentBg;
  const bodyFg = (theme.bodyForegroundColor && String(theme.bodyForegroundColor).trim()) ? theme.bodyForegroundColor : currentText;

  // Apply to document variables
  try {
    document.documentElement.style.setProperty("--bg", bodyBg);
    document.documentElement.style.setProperty("--surface", bodyBg);
    document.documentElement.style.setProperty("--text", bodyFg);

    // Optionally adjust border and muted for contrast if Office provides tokens (example)
    if (theme.bodyBorderColor) {
      document.documentElement.style.setProperty("--border", theme.bodyBorderColor);
    }
    if (theme.bodySubtleColor) {
      document.documentElement.style.setProperty("--muted", theme.bodySubtleColor);
    }
  } catch (e) {
    console.warn("applyOfficeThemeVars: failed to set CSS variables", e);
  }
}

// Register Office theme change handler safely
function registerThemeChangeHandler() {
  try {
    const mailbox = window.Office?.context?.mailbox;
    if (!mailbox || !window.Office.EventType?.OfficeThemeChanged) return;

    mailbox.addHandlerAsync(
      Office.EventType.OfficeThemeChanged,
      function (eventArgs) {
        const theme = eventArgs?.officeTheme;
        applyOfficeThemeVars(theme);
      },
      function (result) {
        if (result && result.status === Office.AsyncResultStatus.Failed) {
          console.error('Failed to register theme change handler:', result.error && result.error.message);
        } else {
          console.debug('Theme change handler registered.');
          // if mailbox exposes current theme, try to apply it immediately
          try {
            const currentTheme = mailbox.officeTheme;
            if (currentTheme) applyOfficeThemeVars(currentTheme);
          } catch { /* ignore */ }
        }
      }
    );
  } catch (e) {
    console.warn('Office theming not available or handler registration failed.', e);
  }
}

/**
 * Logs a message to the console with a [DEBUG] prefix.
 * @param {string} message - The message to log.
 */
function log(message) {
  console.log(`[DEBUG] ${message}`);
}

// A private variable to hold our configuration data.
let configMap = new Map();

// A module-scoped draft string that holds the last-generated HTML draft
let draft = "";

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
  } catch (err) {
    console.error(`Failed to load configuration: ${err.message}`);
    throw err;
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

  if (typeof content === "string" && content.startsWith(fence) && content.endsWith('```')) {
    return content.slice(fence.length, -3);
  }

  return content;
}

/* -------------------------
   DOM helpers and UI updates
   ------------------------- */

/**
 * Shows an error message in the response container (or alert fallback)
 * @param {string} msg - message to show
 */
function showError(msg) {
  const rc = document.getElementById("responseContainer");
  if (rc) {
    rc.textContent = msg;
    rc.classList.add("error");
    rc.setAttribute("role", "alert");
    try { rc.focus(); } catch { /* ignore focus failures */ }
  } else {
    window.alert(msg);
  }
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
  document.getElementById("sentiment").textContent = sentiment ?? "—";
  document.getElementById("urgency").textContent = urgency ?? "—";
  document.getElementById("intention").textContent = intention ?? "—";

  // Update urgency indicator color
  if (urgencyIndicator) {
    urgencyIndicator.classList.remove("urgency-high", "urgency-medium", "urgency-low", "unknown"); // Reset state
    const processedUrgency = String(urgency || "").trim().toLowerCase();

    if (processedUrgency.includes("high")) {
      urgencyIndicator.classList.add("urgency-high");
    } else if (processedUrgency.includes("medium")) {
      urgencyIndicator.classList.add("urgency-medium");
    } else if (processedUrgency.includes("low")) {
      urgencyIndicator.classList.add("urgency-low");
    } else {
      urgencyIndicator.classList.add("unknown");
    }
  }
  return;
}

/* -------------------------
   API call helpers (retry/backoff)
   ------------------------- */

/**
 * Calls the custom API to generate a response.
 * Includes a retry mechanism with exponential backoff for rate limit errors.
 * @param {string} userQuery - The user's query to send to the API.
 * @returns {Promise<string>} A promise that resolves with the generated text.
 */
async function callCustomEndpoint(userQuery) {
  const apiUrl = getVar("customendpoint_url");
  if (!apiUrl) throw new Error("customendpoint_url not configured.");
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

/* -------------------------
   Email helpers
   ------------------------- */

/**
 * Retrieves the email body text for the provided item.
 * @param {Office.Item} item - The mailbox item.
 * @returns {Promise<string>} The plain text body content.
 */
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

/* -------------------------
   Compose helpers and button wiring
   ------------------------- */

/**
 * Returns a sensible suggested subject derived from the item subject.
 * @param {Office.Item} item - The mailbox item.
 * @returns {string} The suggested subject line.
 */
function getSuggestedSubjectFromItem(item) {
  try {
    const subj = item?.subject?.toString ? item.subject.toString() : (item?.subject || "");
    return subj ? `Re: ${subj}` : "Reply";
  } catch {
    return "Reply";
  }
}

/**
 * Opens a compose window using the most appropriate Office.js API available.
 * Tries displayReplyAllForm, displayReplyForm, then displayNewMessageForm as fallback.
 * @param {Office.Item} item - The mailbox item to base the compose on.
 * @param {string} htmlDraft - The HTML content to insert into the compose window.
 * @param {string} fallbackSubject - Optional subject when opening a new message form.
 */
function openComposeWithHtml(item, htmlDraft, fallbackSubject) {
  const htmlBody = typeof htmlDraft === "string" ? htmlDraft : String(htmlDraft || "<p>—</p>");
  try {
    // Preferred: open Reply All compose and inject html
    if (typeof item.displayReplyAllForm === "function") {
      try {
        item.displayReplyAllForm({ htmlBody: htmlBody });
        log("Opened Reply All compose with suggested draft.");
        return;
      } catch (err) {
        log("displayReplyAllForm threw: " + (err?.message || err));
        // fall through to other options
      }
    }

    // Older hosts may only support reply (not replyAll)
    if (typeof item.displayReplyForm === "function") {
      try {
        item.displayReplyForm({ htmlBody: htmlBody });
        log("Opened Reply compose with suggested draft.");
        return;
      } catch (err) {
        log("displayReplyForm threw: " + (err?.message || err));
      }
    }

    // Fallback: open a new message form (recipients won't be prefilled unless added)
    if (typeof Office.context.mailbox.displayNewMessageForm === "function") {
      try {
        Office.context.mailbox.displayNewMessageForm({
          htmlBody: htmlBody,
          subject: fallbackSubject || getSuggestedSubjectFromItem(item),
        });
        log("Opened New Message compose with suggested draft.");
        return;
      } catch (err) {
        log("displayNewMessageForm threw: " + (err?.message || err));
      }
    }

    // No supported compose API available
    showError("Unable to open a draft in this Outlook host. Try replying manually.");
    log("No supported compose APIs available in this host.");
  } catch (err) {
    showError("Failed to open draft. See console for details.");
    console.error("Failed to open draft:", err);
  }
}

/* -------------------------
   DOM wiring for buttons (Quick Reply & Info Button)
   ------------------------- */

document.addEventListener("DOMContentLoaded", () => {
  registerThemeChangeHandler();
  // Ensure Quick Reply button exists and will be enabled by onReady flow
  const btnQuickReply = document.getElementById("btnQuickReply");
  if (btnQuickReply) {
    // The onReady flow will attach the real behavior; keep safe guard here
    btnQuickReply.disabled = true;
  }

  // Info button wiring for help link (THIS WAS THE PREVIOUS FIX FOR THE INFO BUTTON)
  const infoBtn = document.querySelector(".infoBtn");
  if (infoBtn) {
    infoBtn.addEventListener("click", () => {
      try {
        // Assumes 'helpUrl' is configured in config.json
        const helpUrl = getVar("helpUrl");
        if (helpUrl) {
          window.open(helpUrl, "_blank");
          log(`Opened help link: ${helpUrl}`);
        } else {
            const infoWindow = window.open("", "Info", "width=400,height=250,menubar=no,toolbar=no,location=no");
            if (infoWindow) {
              infoWindow.document.write(
                "<html><head><title>About CommsAssist</title></head><body style='font-family:sans-serif;padding:1em;'>" +
                "<h2>CommsAssist</h2>" +
                "<p>Hello! I am your AI assistant for managing email communications.<br>" +
                "I can help you analyze the sentiment, intention, and urgency of incoming emails, " +
                "and even draft responses for you.<br><br>" +
                "Please select an email to get us started!</p>" +
                "</body></html>"
              );
              infoWindow.document.close();
            }
        }
      } catch (err) {
        console.error("Failed to open help link:", err);
        showError("Failed to open help link.");
      }
    });
  }
});

/* -------------------------
   Main Office onReady flow
   ------------------------- */

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
    applyOfficeThemeVars(theme);
  }
  
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
    const quickBtn = document.getElementById("btnQuickReply");
    if (quickBtn) quickBtn.disabled = true;
    // Displays a message to the user
    const rc = document.getElementById("responseContainer");
    if (rc) rc.textContent = "This add-in only works with email messages.";
    return; // Stop execution if no item is available
  }

  log(`Item type: ${item.itemType}`);

  // Checks the mode (Read or Compose) of the mail item
  const isReadMode = item.itemType === Office.MailboxEnums.ItemType.Message;
  const isComposeMode = item.itemType === Office.MailboxEnums.ItemType.MessageCompose;

  if (isReadMode) {
    log("In READ mode.");
    // Shows the read mode UI and enables the Quick Reply button
    document.getElementById("sentimentContent").classList.remove("hidden");
    const quickBtn = document.getElementById("btnQuickReply");
    if (quickBtn) quickBtn.disabled = false;

    // Wait for the email body to be retrieved first
    const emailBody = (await getEmailBody(item)).trim();

    // Now that the 'preview' element is populated, proceed.
    const rc = document.getElementById("responseContainer");
    if (rc) rc.innerHTML = "Analyzing email content, please wait...";

    let name = "";
    if (item.from) {
      name = item.from.displayName || item.from.emailAddress || "";
    }

    try {
      if (customEndpointUrl !== "") {
        // Retrieves additional email metadata asynchronously
        const query = {
          fromEmailAddress: name,
          subject: item.subject ? item.subject : "Unknown",
          body: emailBody,
        };
        const responseData = await callCustomEndpoint(query);
        log(`Call response: ${JSON.stringify(responseData)}`);

        // Safely read expected values with guards
        let sentiment = responseData?.response?.metadata?.email_sentiment ?? "Unknown";
        let urgency = responseData?.response?.metadata?.email_urgency ?? "Unknown";
        let intention = responseData?.response?.metadata?.email_intention ?? "Unknown";
        draft = removeHtmlFences(responseData?.response?.answer?.email_draft ?? "");

        document.getElementById("responseContainer").innerHTML =
          "Analysis complete. Please click 'Quick Reply' to generate a draft.";
        setMetaData(sentiment, urgency, intention);
      } else {
        // No custom endpoint: call Gemini for metadata
        let sentiment = "Unknown";
        let urgency = "Unknown";
        let intention = "Unknown";

        // Calls the Gemini API to analyze sentiment
        try {
          const prompt = `From: ${name}\nBody: ${emailBody}`;
          sentiment = await callGeminiAPI(prompt, sentimentPrompt);
          log(`Sentiment analysis result: ${sentiment}`);
        } catch (err) {
          log(`Error analyzing sentiment: ${err.message}`);
        }
        // Calls the Gemini API to analyze urgency
        try {
          const prompt = `From: ${name}\nBody: ${emailBody}`;
          urgency = await callGeminiAPI(prompt, urgencyPrompt);
          log(`Urgency analysis result: ${urgency}`);
        } catch (err) {
          console.error(`Error analyzing urgency: ${err.message}`);
          log(`Error analyzing urgency: ${err.message}`);
        }
        // Calls the Gemini API to analyze intention
        try {
          const prompt = `From: ${name}\nBody: ${emailBody}`;
          intention = await callGeminiAPI(prompt, intentionPrompt);
          log(`Intention analysis result: ${intention}`);
        } catch (err) {
          console.error(`Error analyzing intention: ${err.message}`);
          log(`Error analyzing intention: ${err.message}`);
        }

        document.getElementById("responseContainer").innerHTML =
          "Analysis complete. Please click 'Quick Reply' to generate a draft.";

        setMetaData(sentiment, urgency, intention);
      }
    } catch (err) {
      log(`Error calling endpoint: ${err.message}`);
      console.error(`Error calling endpoint: ${err.message}`);
      showError("Analysis failed. See console for details.");
    }

    // Event listener for the 'Quick Reply' button (generates and immediately opens a reply)
    const quickReplyBtn = document.getElementById("btnQuickReply");
    if (quickReplyBtn) {
      quickReplyBtn.addEventListener("click", async () => {
        // Disables the button and shows a loading message
        quickReplyBtn.disabled = true;
        const rc2 = document.getElementById("responseContainer");
        if (rc2) rc2.innerHTML = "Generating draft...";

        try {
          if (customEndpointUrl === "") {
            const prompt = `From: ${name}\nBody: ${emailBody}`;
            // Calls the Gemini API to generate the draft
            const generatedText = await callGeminiAPI(prompt, helpdeskPrompt);
            draft = generatedText;
          }
        } catch (err) {
          draft = "<html><body><p>Error generating draft. Please try again.</p></body></html>";
          log(`Error generating draft: ${err.message}`);
          console.error(`Error generating draft: ${err.message}`);
        } finally {
          // Re-enables the button and informs user
          quickReplyBtn.disabled = false;
          if (rc2) rc2.innerHTML = "Draft generated successfully. Opening compose window...";
        }

        // Ensure draft is a string and sanitized for insertion
        const htmlDraft = typeof draft === "string" ? draft : String(draft || "<p></p>");
        const cleanedDraft = removeHtmlFences(htmlDraft);

        // Try to open compose with HTML in a robust order
        openComposeWithHtml(item, cleanedDraft, getSuggestedSubjectFromItem(item));
      });
    }
  } else if (isComposeMode) {
    log("In COMPOSE mode.");
    // Hides the main content and disables buttons for compose mode
    document.getElementById("sentimentContent").classList.add("hidden");
    const quickBtn = document.getElementById("btnQuickReply");
    if (quickBtn) quickBtn.disabled = true;
    const rc = document.getElementById("responseContainer");
    if (rc) rc.innerHTML = "This functionality is not available in compose mode.";
  } else {
    log("In an unsupported mode.");
    // Hides the main content and disables buttons for unsupported modes
    document.getElementById("sentimentContent").classList.add("hidden");
    const quickBtn = document.getElementById("btnQuickReply");
    if (quickBtn) quickBtn.disabled = true;
    const rc = document.getElementById("responseContainer");
    if (rc) rc.textContent = "This add-in only works with email messages.";
  }
});
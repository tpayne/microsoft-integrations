function log(message) {
  console.log(`[DEBUG] ${message}`);
}

Office.onReady(async () => {
  log("Office.js is ready.");
  const item = Office.context?.mailbox?.item;

  if (!item) {
    log("No mail item found. Add-in may be running in an unsupported context.");
    document.getElementById("read_mode_content").classList.add("hidden");
    document.getElementById("btn_quick_reply").disabled = true;
    document.getElementById("btn_generate_draft").disabled = true;
    document.getElementById("response_container").textContent =
      "This add-in only works with email messages.";
    return; // Stop execution if no item is available
  }

  log(`Item type: ${item.itemType}`);

  const isReadMode = item.itemType === Office.MailboxEnums.ItemType.Message;
  const isComposeMode =
    item.itemType === Office.MailboxEnums.ItemType.MessageCompose;

  // Declaring draft variable here to be available to all listeners
  let draft = '';
  
  if (isReadMode) {
    log("In READ mode.");
    document.getElementById("read_mode_content").classList.remove("hidden");
    document.getElementById("btn_quick_reply").disabled = false;
    document.getElementById("btn_generate_draft").disabled = false;

    if (item.subject) {
      document.getElementById("subj").textContent = item.subject;
    }

    if (item.from) {
      document.getElementById("from").textContent =
        item.from.displayName || item.from.emailAddress;
      log(`From retrieved: ${item.from.emailAddress}`);
    }

    if (item.body?.getAsync) {
      item.body.getAsync(Office.CoercionType.Text, (res) => {
        const txt = (res.value || "").trim();
        document.getElementById("preview").textContent = txt;
      });
    }

    document.getElementById("btn_quick_reply").addEventListener("click", async () => {
      try {
        const prompt = `Draft a concise professional IT support response to the following email:
          From: ${document.getElementById("from").textContent}
          Subject: ${document.getElementById("subj").textContent}
          Body: ${document.getElementById("preview").textContent}`;
        const generatedText = await callGeminiAPI(prompt, "Act as a world-class IT support professional. Provide a concise, single-paragraph summary of the key findings.");
        draft = generatedText;
      } catch (error) {
        draft = "Error generating draft. Please try again.";
      }
      Office.context.mailbox.item.displayReplyAllForm({ htmlBody: draft });
    });

    document.getElementById("btn_generate_draft").addEventListener("click", async () => {
      document.getElementById("btn_generate_draft").disabled = true;
      document.getElementById("response_container").textContent = "Generating draft...";

      const prompt = `Draft a concise professional IT support response to the following email:
      From: ${document.getElementById("from").textContent}
      Subject: ${document.getElementById("subj").textContent}
      Body: ${document.getElementById("preview").textContent}`;

      try {
        const generatedText = await callGeminiAPI(prompt, "Act as a world-class IT support professional. Provide a concise, single-paragraph summary of the key findings.");
        draft = generatedText;
        document.getElementById("response_container").textContent = draft;
        log("Draft generated successfully.");
      } catch (error) {
        document.getElementById("response_container").textContent = "Error: Failed to generate a draft.";
        log(`Error generating draft: ${error.message}`);
      } finally {
        document.getElementById("btn_generate_draft").disabled = false;
      }
    });

  } else if (isComposeMode) {
    log("In COMPOSE mode.");
    document.getElementById("read_mode_content").classList.add("hidden");
    document.getElementById("btn_quick_reply").disabled = true;
    document.getElementById("btn_generate_draft").disabled = true;
    document.getElementById("response_container").textContent =
      "This functionality is not available in compose mode.";
  } else {
    log("In an unsupported mode.");
    document.getElementById("read_mode_content").classList.add("hidden");
    document.getElementById("btn_quick_reply").disabled = true;
    document.getElementById("btn_generate_draft").disabled = true;
    document.getElementById("response_container").textContent =
      "This add-in only works with email messages.";
  }
  
  async function callGeminiAPI(userQuery, system_instruction) {
    const apiKey = "<TESTKEYHERE>"; // Replace with your actual API key or use a secure method to store it
    if (!apiKey) {
      throw new Error("GEMINI_API_KEY is not set in environment variables.");
    }
    const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash-preview-05-20:generateContent?key=${apiKey}`;
    const payload = {
      contents: [{ parts: [{ text: userQuery }] }],
      tools: [{ "google_search": {} }],
      systemInstruction: { parts: [{ text: system_instruction }] },
    };

    let attempts = 0;
    const maxAttempts = 5;
    const baseDelay = 1000;

    while (attempts < maxAttempts) {
      try {
        const response = await fetch(apiUrl, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(payload),
        });
        log(`Attempt ${attempts + 1}: API response status is ${response.status}`);

        if (response.status === 429) {
          const delay = baseDelay * Math.pow(2, attempts) + Math.random() * 1000;
          attempts++;
          log(`Rate limit exceeded. Retrying in ${delay}ms.`);
          await new Promise((res) => setTimeout(res, delay));
          continue;
        }

        if (!response.ok) {
          const errorText = await response.text();
          throw new Error(
            `API call failed with status ${response.status}: ${errorText}`
          );
        }

        const result = await response.json();
        const candidate = result.candidates?.[0];

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
          throw new Error(
            `Failed to call Gemini API after ${maxAttempts} attempts.`
          );
        }
      }
    }
  }
});

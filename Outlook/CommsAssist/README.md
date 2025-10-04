# Communications Assistant Outlook Add-in

This project is a sample Outlook add-in that leverages the Gemini API to provide intelligent, AI-powered assistance for drafting email responses. It is designed to help users quickly and professionally respond to emails, especially in IT support scenarios.

## I. Overview

The add-in appears as a task pane in Outlook when reading an email. Its main function is to analyze the email content and provide intelligent assistance using Google's Gemini AI. The tool is now highly robust in a compose/reply context, using a strict internal mechanism to **track the latest edited draft** and prevent general chat queries from inadvertently modifying the email body.

The add-in uses Office.js to read the subject, sender, and body of the current email, creating a detailed prompt for the AI model. It works only in Read Mode and notifies users if accessed in Compose Mode.

**Note** - When deploying this add-in for real, you will need to merge the `pane.css` and `pane.js` files into the main `pane.html` file. You cannot have them separately.

## II. Running Locally

To test the add-in locally, follow these steps:

### Prerequisites

- **Node.js and npm:** Required for running and managing packages.
- **Office Add-in CLI:** Install globally with `npm install -g office-addin-cli`.
- **Running on Unix** This project will only build on Unix systems.

### Steps

1. **Install Dependencies:**
    ```bash
    npm install
    ```

2. **Generate Local Certificates:**
    ```bash
    npm run certs
    ```
    Uses mkcert for trusted HTTPS certificates. You may need to modify the `package.json` for the location of `mkcerts`.

3. **Build the Project:**
    ```bash
    npm run build
    ```
    Copies files to `dist/public`.

4. **Start the Local Server:**
    ```bash
    npm test
    ```
    Serves files at `https://localhost:3000`.

5. **Sideload the Add-in in Outlook:**

    - **Outlook Web:** Go to Settings > Mail > Customize actions > Add-ins > + Add a custom add-in > Add from file...  
      Select `manifest.xml` from the project root.

    - **Outlook Desktop:** Go to File > Options > Customize Ribbon (Windows) or Tools > Get Add-ins (macOS).  
      Upload `manifest.xml` as prompted.

Once loaded, the "Communications Assistant" add-in appears in the Outlook ribbon when reading an email.

![Outlook sample](Images/Chatbot1.png)

![Outlook sample](Images/Chatbot3.png)

![Draft Reply sample](Images/emailDraft.jpeg)

## III. Deploying to Outlook

For production deployment:

1. **Host Files:** Upload `dist/public` contents to a secure HTTPS web server.

2. **Update Manifest:** Replace all `https://localhost:3000/` URLs in `manifest.xml` with your hosted URL (e.g., `https://yourdomain.com/outlook-addin/`).

3. **Deploy Add-in:**
    - **Organizational:** Admin uploads the updated manifest to Microsoft 365 admin center for all users.
    - **Individual:** Users can sideload using the public manifest URL.

## IV. Production Considerations

- **API Key Management:** Do not hardcode the Gemini API key in client-side code. Use a secure backend service to handle API requests.

- **Error Handling & Feedback:** Improve error messages and implement centralized logging.

- **UI/UX Enhancements:** Add loading indicators, improve design, and allow editing drafts within the task pane.

- **Scalability & Performance:** Implement advanced retry, rate-limiting, and monitoring strategies.

- **Manifest & Icons:** Ensure referenced icon files (icon-16.png, icon-32.png, icon-80.png) are created and hosted.

## V. CloudRun Deployment

For CloudRun deployments, you will need to do the following.

- **Build the Image**
  Use Docker to build and test the image

```bash
  docker build . -t outlookserver
  docker run --rm -it -p 8080:3000 outlookserver
  curl -s localhost:8080/pane.html | grep "Communications Assistant" > /dev/null 2>&1
  echo $?
```

  Ensure the returned value is `0`

- **Deploy the Image to CloudRun**
  Deploy the image to CloudRun via

```bash
  gcloud run deploy outlookserver \
    --source . \
    --region us-central1 \
    --platform managed \
    --description "Outlook Server for Comms Assist" \
    --allow-unauthenticated \
    --execution-environment gen2 \
    --ingress all \
    --port 3000
```

# Note
This code has had limited testing, so is not suited for production use

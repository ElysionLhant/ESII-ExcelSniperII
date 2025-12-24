# AI Excel Filler Sidebar

This is an Excel Web Add-in that integrates with LLMs (OpenAI, Volcengine, etc.) to fill Excel tables based on context.

## Features

- **Selection Capture**: Captures the selected range address and a screenshot.
- **Context Materials**: Drag & Drop images, text files (txt, csv, md, json), or PDFs. The add-in reads the file content and sends it to the LLM.
- **Prompt Management**: Save, load, and manage custom prompt presets. Includes default presets for common tasks.
- **Real LLM Integration**: Configurable API endpoint (OpenAI compatible) to generate data.
- **Auto-Expansion**: Automatically expands the Excel range if the LLM generates more rows than selected, preserving formatting.

## How to Run

1. **Prerequisites**: 
   - Node.js installed (to run a local server).
   - An API Key for an OpenAI-compatible service (e.g., OpenAI, Volcengine/Ark).

2. **Start Local Server**:
   Open a terminal in this folder and run:
   ```bash
   npx browser-sync start --server --https --files "." --index "taskpane.html" --port 3000
   ```
   *Note: This starts a secure local server. You might need to accept a certificate warning in your browser.*

3. **Trust the Certificate**:
   Open the URL shown (e.g., `https://localhost:3000/taskpane.html`) in your browser. You will likely see a security warning. Click "Advanced" -> "Proceed to localhost (unsafe)" to trust it temporarily for this session.

4. **Sideload in Excel**:
   - Open Excel (Desktop or Web).
   - Go to the **Insert** tab.
   - Click **Add-ins** (or "Get Add-ins").
   - Click **My Add-ins**.
   - Select **Upload My Add-in** (usually under a "Manage My Add-ins" dropdown or similar).
   - Select the `manifest.xml` file from this folder.

5. **Configuration**:
   - In the sidebar, click **⚙️ Settings**.
   - Enter your **API Base URL** (e.g., `https://api.openai.com/v1` or your Volcengine endpoint).
   - Enter your **API Key**.
   - Enter the **Model Name** (e.g., `gpt-4o` or your Volcengine endpoint ID).

6. **Usage**:
   - Select a range in Excel.
   - Click **Capture Selection**.
   - Drag & drop reference files (PDF/Text) if needed.
   - **Select a Prompt**: Choose a preset from the dropdown or type a custom one. You can manage presets via the "Manage" button.
   - Click **Generate & Fill**.

## Customization

- **Icons**: The manifest points to placeholder icons in `assets/`. You can replace them with real PNGs.

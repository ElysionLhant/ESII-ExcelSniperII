Office.onReady((info) => {
    // UI Event Bindings (Run in both Excel and Browser for testing)
    document.getElementById("capture-btn").onclick = captureSelection;
    document.getElementById("run-btn").onclick = generateAndFill;
    
    // Prompt Management
    document.getElementById("manage-prompts-btn").onclick = togglePromptManager;
    document.getElementById("close-modal-btn").onclick = togglePromptManager;
    document.getElementById("add-prompt-btn").onclick = addNewPrompt;
    document.getElementById("prompt-select").onchange = onPromptSelectChange;
    
    setupDragAndDrop();
    loadPrompts();

    if (info.host === Office.HostType.Excel) {
        // Excel-specific initialization if needed
    }
});

let capturedRangeAddress = null;
let capturedImageBase64 = null;
let uploadedFiles = [];

// --- Prompt Management Logic ---
const DEFAULT_PROMPTS = [
    { title: "General Fill", content: "Fill the table based on the provided image and files." },
    { title: "Invoice Extraction", content: "Extract line items from the invoice image/pdf. Columns: Description, Quantity, Unit Price, Total." },
    { title: "Data Cleanup", content: "Format the data in the image to be consistent and correct any typos." }
];

let savedPrompts = [];

function loadPrompts() {
    const stored = localStorage.getItem("userPrompts");
    if (stored) {
        savedPrompts = JSON.parse(stored);
    } else {
        savedPrompts = [...DEFAULT_PROMPTS];
        savePrompts();
    }
    renderPromptSelect();
    renderPromptList();
}

function savePrompts() {
    localStorage.setItem("userPrompts", JSON.stringify(savedPrompts));
}

function renderPromptSelect() {
    const select = document.getElementById("prompt-select");
    select.innerHTML = '<option value="">-- Select a Preset --</option>';
    savedPrompts.forEach((p, index) => {
        const opt = document.createElement("option");
        opt.value = index;
        opt.text = p.title;
        select.appendChild(opt);
    });
}

function onPromptSelectChange() {
    const index = document.getElementById("prompt-select").value;
    if (index !== "") {
        document.getElementById("prompt-input").value = savedPrompts[index].content;
    }
}

function togglePromptManager() {
    const modal = document.getElementById("prompt-manager");
    modal.style.display = modal.style.display === "none" ? "flex" : "none";
}

function renderPromptList() {
    const list = document.getElementById("prompt-list");
    list.innerHTML = "";
    savedPrompts.forEach((p, index) => {
        const div = document.createElement("div");
        div.className = "prompt-list-item";
        div.innerHTML = `
            <div style="flex:1; margin-right:10px;">
                <strong>${p.title}</strong>
                <div style="font-size:10px; color:#666; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">${p.content}</div>
            </div>
            <button onclick="deletePrompt(${index})" title="Delete">🗑️</button>
        `;
        list.appendChild(div);
    });
}

// Expose to global scope for inline onclick
window.deletePrompt = function(index) {
    if (confirm("Delete this prompt?")) {
        savedPrompts.splice(index, 1);
        savePrompts();
        renderPromptSelect();
        renderPromptList();
    }
};

function addNewPrompt() {
    const title = document.getElementById("new-prompt-title").value;
    const content = document.getElementById("new-prompt-content").value;
    if (title && content) {
        savedPrompts.push({ title, content });
        savePrompts();
        renderPromptSelect();
        renderPromptList();
        document.getElementById("new-prompt-title").value = "";
        document.getElementById("new-prompt-content").value = "";
    }
}
// -------------------------------

async function captureSelection() {
    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.load(["address", "rowCount", "columnCount"]);
            
            // Capture image of the range
            const imageResult = range.getImage();
            
            await context.sync();

            capturedRangeAddress = range.address;
            capturedImageBase64 = imageResult.value;

            // Update UI
            document.getElementById("selection-info").innerText = `Selected: ${range.address} (${range.rowCount} rows x ${range.columnCount} cols)`;
            
            const img = document.getElementById("preview-image");
            img.src = "data:image/png;base64," + imageResult.value;
            img.style.display = "block";
        });
    } catch (error) {
        console.error(error);
        document.getElementById("status-message").innerText = "Error: " + error.message;
    }
}

function setupDragAndDrop() {
    const dropZone = document.getElementById("drop-zone");

    dropZone.addEventListener("dragover", (e) => {
        e.preventDefault();
        dropZone.classList.add("dragover");
    });

    dropZone.addEventListener("dragleave", (e) => {
        e.preventDefault();
        dropZone.classList.remove("dragover");
    });

    dropZone.addEventListener("drop", (e) => {
        e.preventDefault();
        dropZone.classList.remove("dragover");
        handleFiles(e.dataTransfer.files);
    });

    // Paste support
    document.addEventListener("paste", (e) => {
        const items = e.clipboardData.items;
        const files = [];
        for (let i = 0; i < items.length; i++) {
            if (items[i].kind === "file") {
                files.push(items[i].getAsFile());
            }
        }
        if (files.length > 0) handleFiles(files);
    });
}

function handleFiles(files) {
    const fileList = document.getElementById("file-list");
    for (const file of files) {
        uploadedFiles.push(file);
        const div = document.createElement("div");
        div.className = "file-item";
        div.innerText = `${file.name} (${Math.round(file.size/1024)} KB)`;
        fileList.appendChild(div);
    }
}

function toggleSettings() {
    const panel = document.getElementById("settings-panel");
    panel.style.display = panel.style.display === "none" ? "block" : "none";
}

async function generateAndFill() {
    if (!capturedRangeAddress) {
        document.getElementById("status-message").innerText = "Please capture a selection first.";
        return;
    }

    const apiKey = document.getElementById("api-key").value;
    const apiUrl = document.getElementById("api-url").value;
    const modelName = document.getElementById("model-name").value;

    if (!apiKey) {
        document.getElementById("status-message").innerText = "Please enter an API Key in Settings.";
        toggleSettings();
        return;
    }

    document.getElementById("status-message").innerText = "Sending to LLM...";

    try {
        // 1. Prepare Messages
        const userPrompt = document.getElementById("prompt-input").value || "Fill this table based on the image.";
        
        const messages = [
            {
                role: "system",
                content: "You are an Excel data assistant. You will receive an image of an Excel range and a user instruction. Your task is to generate the data to fill the table. \n\nIMPORTANT RESPONSE FORMAT:\n- You MUST return ONLY a raw JSON 2D array (list of lists).\n- Do NOT wrap the output in markdown code blocks (like ```json ... ```).\n- Do NOT include any explanatory text.\n- Example: [[\"Header1\", \"Header2\"], [\"Row1Col1\", \"Row1Col2\"]]"
            },
            {
                role: "user",
                content: [
                    { type: "text", text: userPrompt }
                ]
            }
        ];

        // Add Image if captured
        if (capturedImageBase64) {
            messages[1].content.push({
                type: "image_url",
                image_url: {
                    url: `data:image/png;base64,${capturedImageBase64}`
                }
            });
        }

        // Process uploaded files
        if (uploadedFiles.length > 0) {
            document.getElementById("status-message").innerText = "Reading files...";
            let fileContexts = "";
            
            for (const file of uploadedFiles) {
                try {
                    const content = await readFileContent(file);
                    // Limit content length per file to avoid token limits (simple truncation)
                    const truncatedContent = content.length > 10000 ? content.substring(0, 10000) + "...[truncated]" : content;
                    fileContexts += `\n\n--- File: ${file.name} ---\n${truncatedContent}`;
                } catch (e) {
                    console.error(`Error reading ${file.name}:`, e);
                    fileContexts += `\n\n--- File: ${file.name} ---\n[Error reading file: ${e.message}]`;
                }
            }
            
            if (fileContexts) {
                messages[1].content[0].text += `\n\nUser provided reference materials:${fileContexts}`;
            }
        }

        // 2. Call LLM
        document.getElementById("status-message").innerText = "Sending to LLM...";
        const response = await fetch(`${apiUrl}/chat/completions`, {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
                "Authorization": `Bearer ${apiKey}`
            },
            body: JSON.stringify({
                model: modelName,
                messages: messages,
                max_tokens: 2000,
                temperature: 0.1
            })
        });

        if (!response.ok) {
            throw new Error(`API Error: ${response.status} ${response.statusText}`);
        }

        const data = await response.json();
        const content = data.choices[0].message.content.trim();
        
        // Clean up potential markdown code blocks if the LLM ignores instructions
        const cleanContent = content.replace(/^```json\s*/, "").replace(/^```\s*/, "").replace(/\s*```$/, "");
        
        let tableData;
        try {
            tableData = JSON.parse(cleanContent);
        } catch (e) {
            console.error("JSON Parse Error", content);
            throw new Error("LLM did not return valid JSON array.");
        }

        if (!Array.isArray(tableData) || !Array.isArray(tableData[0])) {
             throw new Error("LLM response is not a 2D array.");
        }

        // 3. Write to Excel
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const originalRange = sheet.getRange(capturedRangeAddress);
            originalRange.load(["rowCount", "columnCount", "rowIndex", "columnIndex"]);
            await context.sync();

            const targetRowCount = tableData.length;
            const currentRowCount = originalRange.rowCount;
            
            // Calculate new range
            let targetRange;
            
            if (targetRowCount > currentRowCount) {
                // Expand
                targetRange = originalRange.getResizedRange(targetRowCount - currentRowCount, 0);
                
                // Copy formatting
                const lastRow = originalRange.getRow(currentRowCount - 1);
                const newRows = targetRange.getRow(currentRowCount).getResizedRange(targetRowCount - currentRowCount - 1, 0);
                newRows.copyFrom(lastRow, Excel.RangeCopyType.formats);
            } else if (targetRowCount < currentRowCount) {
                 // Shrink (Optional: Clear excess rows? For now, just fill what we have)
                 // Actually, let's just resize to the data size so we don't leave old data hanging if we want to be precise,
                 // but usually "filling" implies overwriting. Let's stick to the top-left anchor.
                 targetRange = originalRange.getResizedRange(targetRowCount - currentRowCount, 0);
            } else {
                targetRange = originalRange;
            }

            // Write data
            targetRange.values = tableData;
            targetRange.select();

            await context.sync();
            document.getElementById("status-message").innerText = "Done! Table filled.";
        });

    } catch (error) {
        console.error(error);
        document.getElementById("status-message").innerText = "Error: " + error.message;
    }
}

async function readFileContent(file) {
    if (file.type === "application/pdf") {
        return await readPdfFile(file);
    } else if (file.type.startsWith("text/") || file.name.endsWith(".json") || file.name.endsWith(".csv") || file.name.endsWith(".md") || file.name.endsWith(".txt")) {
        return await readTextFile(file);
    } else {
        return `[File: ${file.name} (Type: ${file.type}) - Content reading not supported for this type. Only Text and PDF are supported.]`;
    }
}

function readTextFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => resolve(e.target.result);
        reader.onerror = (e) => reject(e);
        reader.readAsText(file);
    });
}

async function readPdfFile(file) {
    try {
        const arrayBuffer = await file.arrayBuffer();
        const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
        let fullText = "";
        
        // Limit pages to avoid huge payloads
        const maxPages = Math.min(pdf.numPages, 5); 
        
        for (let i = 1; i <= maxPages; i++) {
            const page = await pdf.getPage(i);
            const textContent = await page.getTextContent();
            const pageText = textContent.items.map(item => item.str).join(" ");
            fullText += `--- Page ${i} ---\n${pageText}\n`;
        }
        
        if (pdf.numPages > maxPages) {
            fullText += `\n... [${pdf.numPages - maxPages} more pages omitted] ...`;
        }
        
        return fullText;
    } catch (e) {
        console.error("PDF Read Error", e);
        throw new Error("Failed to parse PDF. " + e.message);
    }
}
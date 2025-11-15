/* =========================================================================
    Conversation Annotator - Refactored JS (Clean, Modular, No Duplicates)
   ========================================================================= */

(function () {

    /* ---------------------------------------------------------------------
        State
    --------------------------------------------------------------------- */
    const state = {
        currentIndex: 0,
        conversations: [],
        annotations: {},

        buckets: [
            "Bot Response",
            "HVA",
            "AB feature/HVA Related Query",
            "Personalized/Account-Specific Queries",
            "Promo & Freebie Related Queries",
            "Help-page/Direct Customer Service",
            "BP for Non-Profit Organisation Related Query",
            "Personal Prime Related Query",
            "Customer Behavior",
            "Other Queries",
            "Overall Observations"
        ],

        hvaOptions: [
            "3WM (3-Way Match)",
            "Account Authority",
            "Add User",
            "Analytics",
            "ATEP (Amazon Tax Exemption Program)",
            "Business Lists",
            "Business Order Information",
            "Custom Quotes",
            "Guided Buying",
            "PBI",
            "Quantity Discount",
            "Shared Settings",
            "SSO",
            "Subscibe & Save (formerly Recurring Delivery)"
        ]
    };

    /* ---------------------------------------------------------------------
        DOM Elements
    --------------------------------------------------------------------- */
    const el = {
        uploadScreen: document.getElementById("upload-screen"),
        mainInterface: document.getElementById("main-interface"),

        uploadBox: document.getElementById("upload-box"),
        fileInput: document.getElementById("excel-upload"),
        uploadStatus: document.getElementById("upload-status"),

        conversationDisplay: document.getElementById("conversation-display"),
        conversationInfo: document.getElementById("conversation-info"),

        bucketArea: document.getElementById("bucket-area"),

        prevBtn: document.getElementById("prev-btn"),
        nextBtn: document.getElementById("next-btn"),
        saveBtn: document.getElementById("save-btn"),
        downloadBtn: document.getElementById("download-btn"),

        progress: document.getElementById("progress"),
        progressText: document.getElementById("progress-text"),

        statusMessage: document.getElementById("status-message"),
        spinner: document.getElementById("loading-spinner")
    };

    /* ---------------------------------------------------------------------
        Helpers
    --------------------------------------------------------------------- */

    function showSpinner(show) {
        el.spinner.style.display = show ? "flex" : "none";
    }

    function showStatus(msg, type = "info") {
        el.statusMessage.textContent = msg;
        el.statusMessage.className = `status-message alert alert-${type}`;
        el.statusMessage.style.display = "block";
        setTimeout(() => (el.statusMessage.style.display = "none"), 2500);
    }

    function arrayBufferToBinary(buf) {
        const arr = new Uint8Array(buf);
        let str = "";
        for (let i = 0; i < arr.length; i++) str += String.fromCharCode(arr[i]);
        return str;
    }

    /* ---------------------------------------------------------------------
        Bucket UI Rendering
    --------------------------------------------------------------------- */

    function renderBuckets() {
        el.bucketArea.innerHTML = ""; // Prevent duplicates

        state.buckets.forEach(bucket => {

            const isHVA = bucket === "HVA";

            const innerField = isHVA
                ? `
                    <select class="form-select hva-dropdown" name="${bucket}-select">
                        <option value="">Select HVA type...</option>
                        ${state.hvaOptions.map(opt => `<option value="${opt}">${opt}</option>`).join("")}
                    </select>
                `
                : `
                    <textarea name="${bucket}" rows="3" placeholder="Add comments for ${bucket}"></textarea>
                `;

            const html = `
                <div class="bucket" data-bucket="${bucket}">
                    <label class="bucket-label">
                        <input type="checkbox" name="${bucket}">
                        <span>${bucket}</span>
                    </label>

                    <div class="bucket-comment">
                        ${innerField}
                    </div>
                </div>
            `;
            el.bucketArea.insertAdjacentHTML("beforeend", html);
        });
    }

    /* Toggle bucket display */
    el.bucketArea.addEventListener("change", (e) => {
        if (e.target.type !== "checkbox") return;

        const bucketDiv = e.target.closest(".bucket");
        const comment = bucketDiv.querySelector(".bucket-comment");
        const input = bucketDiv.querySelector("textarea, select");

        if (e.target.checked) {
            comment.classList.add("open");
            bucketDiv.classList.add("checked");
            setTimeout(() => input?.focus(), 200);
        } else {
            comment.classList.remove("open");
            bucketDiv.classList.remove("checked");
            if (input) input.value = "";
        }
    });

    /* ---------------------------------------------------------------------
        Excel Handling
    --------------------------------------------------------------------- */

    function readExcel(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();

            reader.onload = e => {
                try {
                    const wb = XLSX.read(new Uint8Array(e.target.result), { type: "array" });
                    const sheet = wb.Sheets[wb.SheetNames[0]];
                    resolve(XLSX.utils.sheet_to_json(sheet));
                } catch (err) {
                    reject(err);
                }
            };

            reader.onerror = reject;
            reader.readAsArrayBuffer(file);
        });
    }

    function processExcelData(rows) {
        const grouped = {};

        rows.forEach(row => {
            if (!grouped[row.Id]) grouped[row.Id] = [];
            grouped[row.Id].push(row);
        });

        state.conversations = Object.values(grouped);
        state.currentIndex = 0;
        state.annotations = {};

        showStatus("File loaded successfully!", "success");
        renderConversation();
    }

    /* ---------------------------------------------------------------------
        Conversation Display
    --------------------------------------------------------------------- */

    function renderConversation() {
        const conv = state.conversations[state.currentIndex];
        const last = conv[conv.length - 1];

        el.conversationInfo.innerHTML = `
            <strong>ID:</strong> ${conv[0].Id} &nbsp; | &nbsp;
            <strong>Feedback:</strong> 
            <span class="badge ${last["Customer Feedback"]?.toLowerCase() === "negative" ? "bg-danger" : "bg-success"}">
                ${last["Customer Feedback"] || "N/A"}
            </span>
        `;

        el.conversationDisplay.innerHTML = conv.map(msg => `
            ${msg.llmGeneratedUserMessage ?
                `<div class="message customer"><div class="message-header">👤 Customer</div>${msg.llmGeneratedUserMessage}</div>`
                : ""}
            ${msg.botMessage ?
                `<div class="message bot"><div class="message-header">🤖 Bot</div>${msg.botMessage}</div>`
                : ""}
        `).join("");

        updateProgress();
        loadAnnotationsForCurrent();
    }

    function updateProgress() {
        const p = ((state.currentIndex + 1) / state.conversations.length) * 100;
        el.progress.style.width = `${p}%`;
        el.progressText.textContent = `${state.currentIndex + 1}/${state.conversations.length} Conversations`;
    }

    /* ---------------------------------------------------------------------
        Annotation Saving + Loading
    --------------------------------------------------------------------- */

    function saveAnnotations() {
        const convId = state.conversations[state.currentIndex][0].Id;

        const selected = {};
        let anyChecked = false;

        state.buckets.forEach(bucket => {
            const checkbox = el.bucketArea.querySelector(`input[name="${bucket}"]`);
            if (checkbox.checked) {
                anyChecked = true;
                selected[bucket] =
                    bucket === "HVA"
                        ? el.bucketArea.querySelector(`select[name="${bucket}-select"]`)?.value || ""
                        : el.bucketArea.querySelector(`textarea[name="${bucket}"]`)?.value?.trim() || "";
            }
        });

        if (!anyChecked) {
            showStatus("Select at least one bucket.", "warning");
            return;
        }

        state.annotations[convId] = selected;
        showStatus("Saved!", "success");
    }

    function loadAnnotationsForCurrent() {
        const convId = state.conversations[state.currentIndex][0].Id;
        const saved = state.annotations[convId] || {};

        state.buckets.forEach(bucket => {
            const bucketDiv = el.bucketArea.querySelector(`[data-bucket="${bucket}"]`);
            const checkbox = bucketDiv.querySelector("input[type=checkbox]");
            const comment = bucketDiv.querySelector(".bucket-comment");
            const input = bucketDiv.querySelector("textarea, select");

            checkbox.checked = false;
            bucketDiv.classList.remove("checked");
            comment.classList.remove("open");
            if (input) input.value = "";

            if (saved[bucket] !== undefined) {
                checkbox.checked = true;
                bucketDiv.classList.add("checked");
                comment.classList.add("open");
                if (input) input.value = saved[bucket];
            }
        });
    }

    /* ---------------------------------------------------------------------
        File Download
    --------------------------------------------------------------------- */

    function downloadExcel() {
        if (Object.keys(state.annotations).length === 0) {
            showStatus("No annotations to download.", "warning");
            return;
        }

        showSpinner(true);

        const output = [];

        state.conversations.forEach(conv => {
            const id = conv[0].Id;
            const ann = state.annotations[id];

            if (!ann) return;

            conv.forEach((msg, i) => {
                const row = {
                    Id: msg.Id,
                    llmGeneratedUserMessage: msg.llmGeneratedUserMessage || "",
                    botMessage: msg.botMessage || "",
                    "Customer Feedback": i === conv.length - 1 ? (msg["Customer Feedback"] || "") : ""
                };

                state.buckets.forEach(b => {
                    row[b] = i === 0 ? ann[b] || "" : "";
                });

                output.push(row);
            });
        });

        const sheet = XLSX.utils.json_to_sheet(output);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, sheet, "Annotations");

        const wbout = XLSX.write(wb, { bookType: "xlsx", type: "binary" });
        const blob = new Blob([arrayBufferToBinary(wbout)], { type: "application/octet-stream" });

        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = `annotated_conversations_${Date.now()}.xlsx`;
        a.click();

        URL.revokeObjectURL(url);
        showSpinner(false);
        showStatus("Downloaded!", "success");
    }

    /* ---------------------------------------------------------------------
        Navigation
    --------------------------------------------------------------------- */

    function nextConversation() {
        if (state.currentIndex >= state.conversations.length - 1)
            return showStatus("Last conversation.", "warning");

        state.currentIndex++;
        renderConversation();
    }

    function prevConversation() {
        if (state.currentIndex === 0)
            return showStatus("First conversation.", "warning");

        state.currentIndex--;
        renderConversation();
    }

    /* ---------------------------------------------------------------------
        File Upload Events
    --------------------------------------------------------------------- */

    el.uploadBox.addEventListener("click", () => el.fileInput.click());

    el.fileInput.addEventListener("change", async e => {
        const file = e.target.files[0];
        if (!file || !file.name.endsWith(".xlsx")) {
            return showStatus("Upload a valid .xlsx file.", "danger");
        }

        try {
            showSpinner(true);
            const data = await readExcel(file);
            processExcelData(data);

            el.uploadScreen.style.display = "none";
            el.mainInterface.style.display = "flex";
        } catch (err) {
            console.error(err);
            showStatus("Error loading Excel file.", "danger");
        } finally {
            showSpinner(false);
        }
    });

    /* Drag-drop */
    el.uploadBox.addEventListener("dragover", e => {
        e.preventDefault();
        el.uploadBox.classList.add("dragover");
    });

    el.uploadBox.addEventListener("dragleave", () => {
        el.uploadBox.classList.remove("dragover");
    });

    el.uploadBox.addEventListener("drop", e => {
        e.preventDefault();
        el.uploadBox.classList.remove("dragover");

        const file = e.dataTransfer.files[0];
        if (file && file.name.endsWith(".xlsx")) {
            el.fileInput.files = e.dataTransfer.files;
            el.fileInput.dispatchEvent(new Event("change"));
        } else {
            showStatus("Upload a valid .xlsx file.", "danger");
        }
    });

    /* ---------------------------------------------------------------------
        Button Events
    --------------------------------------------------------------------- */

    el.nextBtn.addEventListener("click", nextConversation);
    el.prevBtn.addEventListener("click", prevConversation);
    el.saveBtn.addEventListener("click", saveAnnotations);
    el.downloadBtn.addEventListener("click", downloadExcel);

    /* Keyboard shortcuts */
    document.addEventListener("keydown", e => {
        if (el.mainInterface.style.display === "none") return;

        if (e.key === "ArrowRight") nextConversation();
        if (e.key === "ArrowLeft") prevConversation();
        if (e.key === "s" && (e.ctrlKey || e.metaKey)) {
            e.preventDefault();
            saveAnnotations();
        }
    });

    /* ---------------------------------------------------------------------
        Init
    --------------------------------------------------------------------- */
    renderBuckets();

})();

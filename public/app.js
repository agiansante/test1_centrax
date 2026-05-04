const form = document.querySelector("#chat-form");
const promptInput = document.querySelector("#prompt");
const messagesEl = document.querySelector("#messages");
const statusEl = document.querySelector("#status");
const fileInput = document.querySelector("#file-input");
const attachButton = document.querySelector("#attach-button");
const attachmentBar = document.querySelector("#attachment-bar");

const messages = [];
const selectedFiles = [];
const uploadedDocuments = [];
const maxTotalFileBytes = 50 * 1024 * 1024;
const textExtensions = new Set([
  "txt",
  "md",
  "markdown",
  "csv",
  "tsv",
  "json",
  "jsonl",
  "xml",
  "html",
  "css",
  "js",
  "jsx",
  "ts",
  "tsx",
  "py",
  "java",
  "cs",
  "cpp",
  "c",
  "h",
  "sql",
  "yaml",
  "yml",
  "log"
]);

function setStatus(text) {
  statusEl.textContent = text;
}

function addMessage(role, content) {
  const article = document.createElement("article");
  article.className = `message ${role}`;

  const paragraph = document.createElement("p");
  paragraph.textContent = content;

  article.append(paragraph);
  messagesEl.append(article);
  messagesEl.scrollTop = messagesEl.scrollHeight;
}

function formatBytes(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
}

function renderAttachments() {
  attachmentBar.replaceChildren();
  attachmentBar.hidden = selectedFiles.length === 0 && uploadedDocuments.length === 0;

  selectedFiles.forEach((file, index) => {
    const chip = document.createElement("div");
    chip.className = "attachment-chip";

    const label = document.createElement("span");
    label.textContent = `Da caricare: ${file.name} (${formatBytes(file.size)})`;
    label.title = label.textContent;

    const removeButton = document.createElement("button");
    removeButton.type = "button";
    removeButton.textContent = "x";
    removeButton.ariaLabel = `Rimuovi ${file.name}`;
    removeButton.addEventListener("click", () => {
      selectedFiles.splice(index, 1);
      renderAttachments();
    });

    chip.append(label, removeButton);
    attachmentBar.append(chip);
  });

  uploadedDocuments.forEach((uploadedDocument, index) => {
    const chip = document.createElement("div");
    chip.className = "attachment-chip active";

    const label = document.createElement("span");
    const mode = uploadedDocument.text
      ? "indicizzato"
      : isWordFile(uploadedDocument)
        ? "lettura Word"
        : "allegato originale";
    label.textContent = `${uploadedDocument.name} (${mode})`;
    label.title = label.textContent;

    const removeButton = document.createElement("button");
    removeButton.type = "button";
    removeButton.textContent = "x";
    removeButton.ariaLabel = `Rimuovi ${uploadedDocument.name}`;
    removeButton.addEventListener("click", () => {
      uploadedDocuments.splice(index, 1);
      renderAttachments();
    });

    chip.append(label, removeButton);
    attachmentBar.append(chip);
  });
}

function isTextLikeFile(file) {
  const extension = file.name.split(".").pop()?.toLowerCase() || "";

  if (isWordFile(file) || extension === "xlsx" || extension === "pptx") {
    return false;
  }

  return file.type.startsWith("text/")
    || file.type.includes("json")
    || file.type.includes("xml")
    || file.type.includes("csv")
    || file.type.includes("javascript")
    || textExtensions.has(extension);
}

function isWordFile(file) {
  return file.name.toLowerCase().endsWith(".docx")
    || file.type === "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
}

function readFileAsText(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.addEventListener("load", () => resolve(String(reader.result || "")));
    reader.addEventListener("error", () => reject(new Error(`Non riesco a leggere ${file.name}.`)));
    reader.readAsText(file);
  });
}

function readFileAsDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.addEventListener("load", () => resolve(reader.result));
    reader.addEventListener("error", () => reject(new Error(`Non riesco a leggere ${file.name}.`)));
    reader.readAsDataURL(file);
  });
}

function createId() {
  return crypto.randomUUID?.() || `${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

async function serializeSelectedFiles() {
  const totalBytes = [...uploadedDocuments, ...selectedFiles].reduce((sum, file) => sum + file.size, 0);

  if (totalBytes > maxTotalFileBytes) {
    throw new Error("I file caricati superano il limite totale di 50 MB.");
  }

  return Promise.all(selectedFiles.map(async (file) => {
    const text = isTextLikeFile(file) ? await readFileAsText(file) : "";

    return {
      id: createId(),
      name: file.name,
      type: file.type || "application/octet-stream",
      size: file.size,
      text,
      dataUrl: text ? "" : await readFileAsDataUrl(file)
    };
  }));
}

function resizeInput() {
  promptInput.style.height = "auto";
  promptInput.style.height = `${promptInput.scrollHeight}px`;
}

async function sendMessage(content, newDocuments) {
  const userMessage = { role: "user", content };
  messages.push(userMessage);

  const fileSummary = newDocuments.length
    ? `\n\nFile caricati: ${newDocuments.map((file) => file.name).join(", ")}`
    : "";

  addMessage("user", `${content || "Analizza i file caricati."}${fileSummary}`);
  setStatus("Sto pensando...");

  const response = await fetch("/api/chat", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ messages, documents: uploadedDocuments })
  });

  const data = await response.json();

  if (!response.ok) {
    throw new Error(data.error || "Errore durante la risposta.");
  }

  messages.push({ role: "assistant", content: data.reply });
  addMessage("assistant", data.reply);
}

form.addEventListener("submit", async (event) => {
  event.preventDefault();

  const content = promptInput.value.trim();
  if (!content && selectedFiles.length === 0) return;

  promptInput.value = "";
  resizeInput();
  promptInput.disabled = true;
  attachButton.disabled = true;
  form.querySelector("button[type='submit']").disabled = true;

  try {
    const newDocuments = await serializeSelectedFiles();
    uploadedDocuments.push(...newDocuments);
    selectedFiles.length = 0;
    renderAttachments();

    await sendMessage(content, newDocuments);
    setStatus("Pronta");
  } catch (error) {
    addMessage("assistant", error.message);
    setStatus("Errore");
  } finally {
    promptInput.disabled = false;
    attachButton.disabled = false;
    form.querySelector("button[type='submit']").disabled = false;
    promptInput.focus();
  }
});

attachButton.addEventListener("click", () => fileInput.click());
fileInput.addEventListener("change", () => {
  selectedFiles.push(...fileInput.files);
  fileInput.value = "";
  renderAttachments();
});

promptInput.addEventListener("input", resizeInput);
promptInput.addEventListener("keydown", (event) => {
  if (event.key === "Enter" && !event.shiftKey) {
    event.preventDefault();
    form.requestSubmit();
  }
});

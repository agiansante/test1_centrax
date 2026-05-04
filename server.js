import { createServer } from "node:http";
import { readFile } from "node:fs/promises";
import { extname, join, normalize, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { inflateRawSync } from "node:zlib";

const __dirname = fileURLToPath(new URL(".", import.meta.url));
const publicDir = join(__dirname, "public");
const port = Number(process.env.PORT || 3000);
const maxFileBytes = 50 * 1024 * 1024;
const maxTotalFileBytes = 50 * 1024 * 1024;
const chunkSize = 2800;
const chunkOverlap = 320;
const maxContextChunks = 10;
const maxContextCharacters = 28000;

const mimeTypes = {
  ".html": "text/html; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8"
};

async function loadLocalEnv() {
  try {
    const envFile = await readFile(join(__dirname, ".env"), "utf8");

    for (const line of envFile.split(/\r?\n/)) {
      const trimmed = line.trim();
      if (!trimmed || trimmed.startsWith("#")) continue;

      const separatorIndex = trimmed.indexOf("=");
      if (separatorIndex === -1) continue;

      const key = trimmed.slice(0, separatorIndex).trim();
      const value = trimmed.slice(separatorIndex + 1).trim().replace(/^["']|["']$/g, "");

      if (key && process.env[key] === undefined) {
        process.env[key] = value;
      }
    }
  } catch {
    // The .env file is optional for local experiments.
  }
}

function sendJson(res, statusCode, payload) {
  res.writeHead(statusCode, { "Content-Type": "application/json; charset=utf-8" });
  res.end(JSON.stringify(payload));
}

async function readJsonBody(req) {
  const chunks = [];

  for await (const chunk of req) {
    chunks.push(chunk);
  }

  const rawBody = Buffer.concat(chunks).toString("utf8");
  return rawBody ? JSON.parse(rawBody) : {};
}

function toResponseText(data) {
  if (typeof data.output_text === "string" && data.output_text.trim()) {
    return data.output_text;
  }

  return (data.output || [])
    .flatMap((item) => item.content || [])
    .map((content) => content.text)
    .filter(Boolean)
    .join("\n")
    .trim();
}

function dataUrlToBuffer(dataUrl) {
  const base64Marker = ";base64,";
  const markerIndex = dataUrl.indexOf(base64Marker);

  if (markerIndex === -1) {
    return Buffer.alloc(0);
  }

  return Buffer.from(dataUrl.slice(markerIndex + base64Marker.length), "base64");
}

function isDocxFile(file) {
  return file.name.toLowerCase().endsWith(".docx")
    || file.type === "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
}

function looksLikeBinaryZipText(text) {
  return text.startsWith("PK") && /word\/document\.xml|word\/_rels|docProps\//.test(text);
}

function decodeXmlText(xml) {
  return xml
    .replace(/<w:tab\/>/g, "\t")
    .replace(/<w:br\/>|<w:cr\/>/g, "\n")
    .replace(/<\/w:p>/g, "\n")
    .replace(/<\/w:tr>/g, "\n")
    .replace(/<\/w:tc>/g, "\t")
    .replace(/<[^>]+>/g, "")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, "\"")
    .replace(/&apos;/g, "'")
    .replace(/\n{3,}/g, "\n\n")
    .replace(/[ \t]{2,}/g, " ")
    .trim();
}

function readZipEntries(buffer) {
  const entries = new Map();
  let offset = 0;

  while (offset < buffer.length - 30) {
    const signature = buffer.readUInt32LE(offset);
    if (signature !== 0x04034b50) {
      offset += 1;
      continue;
    }

    const method = buffer.readUInt16LE(offset + 8);
    const compressedSize = buffer.readUInt32LE(offset + 18);
    const fileNameLength = buffer.readUInt16LE(offset + 26);
    const extraLength = buffer.readUInt16LE(offset + 28);
    const nameStart = offset + 30;
    const nameEnd = nameStart + fileNameLength;
    const dataStart = nameEnd + extraLength;
    const dataEnd = dataStart + compressedSize;

    if (dataEnd > buffer.length) break;

    const name = buffer.toString("utf8", nameStart, nameEnd);
    const compressed = buffer.subarray(dataStart, dataEnd);

    if (method === 0) {
      entries.set(name, compressed);
    } else if (method === 8) {
      entries.set(name, inflateRawSync(compressed));
    }

    offset = dataEnd;
  }

  return entries;
}

function extractDocxText(dataUrl) {
  const entries = readZipEntries(dataUrlToBuffer(dataUrl));
  const xmlNames = [...entries.keys()].filter((name) => (
    name === "word/document.xml"
    || /^word\/(header|footer|footnotes|endnotes)\d*\.xml$/.test(name)
  ));

  return xmlNames
    .map((name) => decodeXmlText(entries.get(name).toString("utf8")))
    .filter(Boolean)
    .join("\n\n")
    .trim();
}

function normalizeAttachments(attachments) {
  if (!Array.isArray(attachments)) return [];

  let totalBytes = 0;

  return attachments.map((attachment) => {
    const name = String(attachment.name || "allegato").slice(0, 180);
    const type = String(attachment.type || "application/octet-stream");
    const size = Number(attachment.size || 0);
    const dataUrl = String(attachment.dataUrl || "");
    const text = typeof attachment.text === "string" ? attachment.text : "";

    if (!text && (!dataUrl.startsWith("data:") || !dataUrl.includes(";base64,"))) {
      throw new Error(`Il file "${name}" non ha un formato valido.`);
    }

    if (size > maxFileBytes) {
      throw new Error(`Il file "${name}" supera il limite di 50 MB.`);
    }

    totalBytes += size;
    if (totalBytes > maxTotalFileBytes) {
      throw new Error("Gli allegati superano il limite totale di 50 MB.");
    }

    let extractedText = text;

    if (isDocxFile({ name, type }) && dataUrl && (!extractedText || looksLikeBinaryZipText(extractedText))) {
      extractedText = extractDocxText(dataUrl);
    }

    return { name, type, size, dataUrl, text: extractedText };
  });
}

function chunkDocument(document) {
  const text = String(document.text || "").replace(/\r\n/g, "\n").trim();
  if (!text) return [];

  const chunks = [];
  let start = 0;

  while (start < text.length) {
    const end = Math.min(start + chunkSize, text.length);
    chunks.push({
      fileName: document.name,
      index: chunks.length + 1,
      text: text.slice(start, end)
    });

    if (end === text.length) break;
    start = Math.max(0, end - chunkOverlap);
  }

  return chunks;
}

function tokenize(text) {
  return String(text || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .match(/[a-z0-9]{3,}/g) || [];
}

function scoreChunk(chunk, terms) {
  if (terms.length === 0) return 0;

  const haystack = chunk.text.toLowerCase();
  return terms.reduce((score, term) => score + (haystack.includes(term) ? 1 : 0), 0);
}

function buildDocumentContext(documents, question) {
  if (!Array.isArray(documents) || documents.length === 0) {
    return { contextText: "", rawFiles: [] };
  }

  const normalizedDocuments = normalizeAttachments(documents);
  const textDocuments = normalizedDocuments.filter((document) => document.text.trim());
  const rawFiles = normalizedDocuments.filter((document) => !document.text.trim() && document.dataUrl);
  const chunks = textDocuments.flatMap(chunkDocument);

  if (chunks.length === 0) {
    return { contextText: "", rawFiles };
  }

  const terms = tokenize(question);
  const rankedChunks = chunks
    .map((chunk) => ({ ...chunk, score: scoreChunk(chunk, terms) }))
    .sort((a, b) => b.score - a.score || a.fileName.localeCompare(b.fileName) || a.index - b.index);

  const selectedChunks = rankedChunks.some((chunk) => chunk.score > 0)
    ? rankedChunks.filter((chunk) => chunk.score > 0).slice(0, maxContextChunks)
    : rankedChunks.slice(0, maxContextChunks);

  let usedCharacters = 0;
  const contextParts = [];

  for (const chunk of selectedChunks) {
    const header = `[${chunk.fileName} - chunk ${chunk.index}]`;
    const part = `${header}\n${chunk.text}`;

    if (usedCharacters + part.length > maxContextCharacters) break;

    contextParts.push(part);
    usedCharacters += part.length;
  }

  return {
    contextText: contextParts.length
      ? `Documenti caricati dall'utente. Usa questi estratti come fonte principale quando rispondi.\n\n${contextParts.join("\n\n---\n\n")}`
      : "",
    rawFiles
  };
}

function toModelMessage(message) {
  const role = message.role === "assistant" ? "assistant" : "user";
  const text = String(message.content || "");

  if (role === "assistant") {
    return { role, content: text };
  }

  const content = [];
  const attachments = normalizeAttachments(message.attachments);

  for (const attachment of attachments) {
    if (attachment.type.startsWith("image/")) {
      content.push({
        type: "input_image",
        image_url: attachment.dataUrl,
        detail: "auto"
      });
    } else {
      content.push({
        type: "input_file",
        filename: attachment.name,
        file_data: attachment.dataUrl
      });
    }
  }

  const textAttachments = attachments.filter((attachment) => attachment.text.trim());
  if (textAttachments.length > 0) {
    const loadedFiles = textAttachments
      .map((attachment) => `- ${attachment.name} (${attachment.text.length} caratteri letti)`)
      .join("\n");

    content.push({
      type: "input_text",
      text: `File caricati e indicizzati per la chat:\n${loadedFiles}`
    });
  }

  content.push({
    type: "input_text",
    text: text || "Analizza gli allegati caricati."
  });

  return { role, content };
}

async function handleChat(req, res) {
  if (!process.env.OPENAI_API_KEY) {
    return sendJson(res, 500, {
      error: "Manca OPENAI_API_KEY. Impostala nel terminale prima di avviare il server."
    });
  }

  try {
    const { messages, documents } = await readJsonBody(req);

    if (!Array.isArray(messages) || messages.length === 0) {
      return sendJson(res, 400, { error: "Invia almeno un messaggio." });
    }

    const latestQuestion = [...messages].reverse().find((message) => message.role !== "assistant")?.content || "";
    const { contextText, rawFiles } = buildDocumentContext(documents, latestQuestion);
    const input = [];

    if (contextText || rawFiles.length > 0) {
      const contextContent = [];

      for (const file of rawFiles) {
        if (file.type.startsWith("image/")) {
          contextContent.push({
            type: "input_image",
            image_url: file.dataUrl,
            detail: "auto"
          });
        } else {
          contextContent.push({
            type: "input_file",
            filename: file.name,
            file_data: file.dataUrl
          });
        }
      }

      if (contextText) {
        contextContent.push({
          type: "input_text",
          text: contextText
        });
      }

      input.push({ role: "user", content: contextContent });
    }

    input.push(...messages.map(toModelMessage));

    const apiResponse = await fetch("https://api.openai.com/v1/responses", {
      method: "POST",
      headers: {
        "Authorization": `Bearer ${process.env.OPENAI_API_KEY}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        model: process.env.OPENAI_MODEL || "gpt-5.2",
        instructions: "Rispondi in italiano in modo chiaro, utile e conciso. Quando sono presenti documenti caricati, usali come contesto principale e segnala se l'informazione non è nel file.",
        input
      })
    });

    const data = await apiResponse.json();

    if (!apiResponse.ok) {
      const message = data?.error?.message || "Errore nella chiamata al modello.";
      return sendJson(res, apiResponse.status, { error: message });
    }

    return sendJson(res, 200, {
      reply: toResponseText(data) || "Non ho ricevuto testo dal modello."
    });
  } catch (error) {
    return sendJson(res, 500, {
      error: error instanceof Error ? error.message : "Errore inatteso."
    });
  }
}

async function serveStatic(req, res) {
  const url = new URL(req.url || "/", `http://${req.headers.host}`);
  const requestedPath = url.pathname === "/" ? "/index.html" : url.pathname;
  const safePath = normalize(decodeURIComponent(requestedPath)).replace(/^[\/\\]+/, "");
  const filePath = resolve(publicDir, safePath);

  if (!filePath.startsWith(resolve(publicDir))) {
    return sendJson(res, 403, { error: "Percorso non permesso." });
  }

  try {
    const file = await readFile(filePath);
    const contentType = mimeTypes[extname(filePath)] || "application/octet-stream";
    res.writeHead(200, { "Content-Type": contentType });
    res.end(file);
  } catch {
    sendJson(res, 404, { error: "File non trovato." });
  }
}

await loadLocalEnv();

const server = createServer(async (req, res) => {
  if (req.method === "POST" && req.url === "/api/chat") {
    return handleChat(req, res);
  }

  if (req.method === "GET" || req.method === "HEAD") {
    return serveStatic(req, res);
  }

  return sendJson(res, 405, { error: "Metodo non supportato." });
});

server.listen(port, () => {
  console.log(`Chat LLM pronta su http://localhost:${port}`);
});

const STORAGE_KEYS = {
await context.sync();
return range.text;
});
}


async function replaceSelectionText(text) {
return Word.run(async (context) => {
const range = context.document.getSelection();
range.insertText(text, Word.InsertLocation.replace);
await context.sync();
});
}


async function callBackend(payload) {
const res = await fetch(CONFIG.API_ENDPOINT, {
method: "POST",
headers: { "Content-Type": "application/json" },
body: JSON.stringify(payload)
});
const data = await res.json();
if (!data.success) throw new Error(data.error || "Unknown error");
return data.output;
}


async function onGenerate() {
setStatus("Processing...");
const provider = document.getElementById("provider").value;
const apiKey = document.getElementById("apiKey").value;
const mode = document.getElementById("mode").value;
const prompt = document.getElementById("prompt").value;


if (!apiKey) return setStatus("API key required.");
saveSettings();


let selectionText = "";
if (mode !== "generate") selectionText = await getSelectionText();


const output = await callBackend({ provider, apiKey, mode, prompt, selectionText });
await replaceSelectionText(output);
setStatus("Inserted.");
}


async function onLoadSelection() {
const selection = await getSelectionText();
document.getElementById("prompt").value = selection;
setStatus("Selection loaded.");
}


Office.onReady(() => {
document.getElementById("btnGenerate").onclick = onGenerate;
document.getElementById("btnUseSelection").onclick = onLoadSelection;
loadSettings();
setStatus("Ready.");
});

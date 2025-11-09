
// Health check
async function checkHealth() {
  const statusEl = document.getElementById("health-status");
  try {
    const res = await fetch(`/api/health`);
    if (res.ok) {
      statusEl.textContent = "✅ Backend Online v-1.0.1";
      statusEl.className = "status ok";
    } else {
      statusEl.textContent = "❌ Backend Errors";
      statusEl.className = "status down";
    }
  } catch (err) {
    statusEl.textContent = "❌ Backend Offline";
    statusEl.className = "status down";
  }
}

checkHealth();
setInterval(checkHealth, 10000);

// Drag & drop for DOCX
const dropArea = document.getElementById("drop-area");
const fileInput = document.getElementById("template");

dropArea.addEventListener("click", () => fileInput.click());

dropArea.addEventListener("dragover", e => {
  e.preventDefault();
  dropArea.classList.add("dragover");
});

dropArea.addEventListener("dragleave", () => dropArea.classList.remove("dragover"));

dropArea.addEventListener("drop", e => {
  e.preventDefault();
  dropArea.classList.remove("dragover");
  if (e.dataTransfer.files.length > 0) {
    fileInput.files = e.dataTransfer.files;
  }
});

// Load sample JSON
document.getElementById("loadSampleBtn").addEventListener("click", async () => {
  try {
    const res = await fetch("/samples/test_data.json");
    const jsonText = await res.text();
    document.getElementById("jsonData").value = jsonText;
  } catch (err) {
    alert("Failed to load sample JSON: " + err.message);
  }
});

// Generate DOCX
document.getElementById("generateBtn").addEventListener("click", async () => {
  if (!fileInput.files.length) return alert("Please choose a DOCX template.");

  const jsonText = document.getElementById("jsonData").value;
  let parsedJson;
  try {
    parsedJson = JSON.parse(jsonText);
  } catch (err) {
    return alert("Invalid JSON: " + err.message);
  }

  const formData = new FormData();
  formData.append("template", fileInput.files[0]);
  formData.append("data", JSON.stringify(parsedJson));

  try {
    const response = await fetch(`/api/generate-direct`, {
      method: "POST",
      body: formData
    });

    if (!response.ok) {
      const text = await response.text();
      throw new Error(text || "Failed to generate DOCX");
    }

    const blob = await response.blob();
    const url = window.URL.createObjectURL(blob);

    const a = document.createElement("a");
    a.href = url;
    a.download = "generated.docx";
    document.body.appendChild(a);
    a.click();
    a.remove();
  } catch (err) {
    alert("Error: " + err.message);
  }
});

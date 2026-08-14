const dropzone = document.getElementById("dropzone");
const fileInput = document.getElementById("file-input");
const results = document.getElementById("results");

dropzone.addEventListener("click", () => fileInput.click());

fileInput.addEventListener("change", () => {
  if (fileInput.files.length) {
    handleFiles(fileInput.files);
    fileInput.value = "";
  }
});

["dragenter", "dragover"].forEach((evt) =>
  dropzone.addEventListener(evt, (e) => {
    e.preventDefault();
    dropzone.classList.add("dragover");
  })
);

["dragleave", "drop"].forEach((evt) =>
  dropzone.addEventListener(evt, (e) => {
    e.preventDefault();
    dropzone.classList.remove("dragover");
  })
);

dropzone.addEventListener("drop", (e) => {
  const files = e.dataTransfer.files;
  if (files.length) handleFiles(files);
});

function handleFiles(fileList) {
  const files = Array.from(fileList);
  const cards = new Map();

  files.forEach((file) => {
    const card = document.createElement("div");
    card.className = "result-card pending";
    card.innerHTML = `<div class="filename">${escapeHtml(file.name)}</div><div class="status-line">Converting…</div>`;
    results.prepend(card);
    cards.set(file.name, card);
  });

  const formData = new FormData();
  files.forEach((file) => formData.append("files", file));

  fetch("/convert", { method: "POST", body: formData })
    .then((res) => res.json())
    .then((data) => {
      (data.results || []).forEach((result) => {
        const card = cards.get(result.filename) || appendNewCard(result.filename);
        renderResult(card, result);
      });
    })
    .catch((err) => {
      cards.forEach((card) => {
        card.className = "result-card error";
        card.querySelector(".status-line").textContent = "Upload failed: " + err;
      });
    });
}

function appendNewCard(filename) {
  const card = document.createElement("div");
  card.className = "result-card pending";
  card.innerHTML = `<div class="filename">${escapeHtml(filename || "File")}</div>`;
  results.prepend(card);
  return card;
}

function renderResult(card, result) {
  if (result.success) {
    card.className = "result-card success";
    card.innerHTML = `
      <div class="filename">${escapeHtml(result.filename)}</div>
      <div class="status-line">Converted successfully.</div>
      <a class="download-link" href="${result.download_url}">Download ${escapeHtml(result.output_filename)}</a>
    `;
  } else {
    card.className = "result-card error";
    card.innerHTML = `
      <div class="filename">${escapeHtml(result.filename || "File")}</div>
      <div class="status-line">${escapeHtml(result.error || "Conversion failed.")}</div>
    `;
  }
}

function escapeHtml(str) {
  const div = document.createElement("div");
  div.textContent = str == null ? "" : str;
  return div.innerHTML;
}

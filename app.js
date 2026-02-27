/* JAGD Daily Maintenance Checklists (Mobile Webform)
   - Mobile-first UI
   - Saves into a clean, non-fillable PDF template (no Acrobat JS prompts / no hidden fields)
   Template: maintenance_template.pdf (based on "JAGD Daily Equipment Inspection Checklist.pdf")
*/

const PROJECT_OPTIONS = [
  "  ",
  "69th St. Transfer Bridge",
  "BA-2024-RE-102-CM Mid-Hudson Bridge",
  "BRX9579 - Boston Road Bridge",
  "BW96 & VN12 - Whitestone Hellman Platforms",
  "C35311 - Dyre Ave. Line",
  "D214898 - TANE22-29 Restani T&M",
  "D264324 - Westchester County Field Metalizing",
  "D264965 - Highway bridge repair W&W",
  "D265046 - Highway bridge repair W&W",
  "D265307 - WO03",
  "D265343 - Bove W&W 2",
  "Devon Bridge",
  "DMB-25-01",
  "FCC 2056",
  "Gold Star Memorial Bridge",
  "Governors Island",
  "Grand Concourse",
  "GW 244.289 Lemoine Ave",
  "GWB Cables",
  "HB1070MD - Macombs Dam Bridge",
  "HBKBQE - NYCDOT Bove",
  "K7279 & K6176 Gordie Howe",
  "Park Avenue",
  "Pulaski 8B",
  "QBB-2017",
  "RK90",
  "Sandy Relief",
  "VN81X",
  "VN-84B - Verrazzano Bridge Ramps Brooklyn",
  "Warehouse"
];

const CONFIG = {
  templateFile: "maintenance_template.pdf",
  pageSize: { w: 612, h: 792 },

  // Checklist table cell geometry (measured from the clean template)
  table: {
    // Row centers (distance from TOP of page, in PDF points)
    rowCentersTop: [125.5, 143.5, 161.5, 179.5, 197.5, 215.5, 233.5, 251.5, 269.5, 287.5],
    okCenterX: 318.5,
    needsCenterX: 376.0,
    commentX: 415.0,
    commentW: 160.0
  },

  // Photo slots (4 per page). Coordinates are PDF points (bottom-left origin).
  photos: [
    { // slot 1 (top-left)
      header: { x: 36.5, y: 454.0, w: 269.0, h: 17.0 },
      box:    { x: 36.5, y: 334.0, w: 269.0, h: 119.0 }
    },
    { // slot 2 (top-right)
      header: { x: 306.5, y: 454.0, w: 269.0, h: 17.0 },
      box:    { x: 306.5, y: 334.0, w: 269.0, h: 119.0 }
    },
    { // slot 3 (bottom-left)
      header: { x: 36.5, y: 316.0, w: 269.0, h: 17.0 },
      box:    { x: 36.5, y: 196.0, w: 269.0, h: 119.0 }
    },
    { // slot 4 (bottom-right)
      header: { x: 306.5, y: 316.0, w: 269.0, h: 17.0 },
      box:    { x: 306.5, y: 196.0, w: 269.0, h: 119.0 }
    }
  ],

  // Text placement for top/bottom fields (PDF points)
  fields: {
    project: { x: 112, y: 702, w: 300, size: 12 },

    additionalComments: { x: 78, yTop: 152, w: 456, size: 10, lineHeight: 12, maxLines: 5 },

    name: { x: 70, y: 44, w: 220, size: 11 },
    date: { x: 515, y: 44, w: 90, size: 11 },

    signature: { x: 300, y: 28, w: 190, h: 55 }
  },

  pages: [
    {
      title: "Air Compressor – Daily Inspection Checklist",
      items: [
        "Engine oil level correct",
        "Coolant level correct",
        "Fuel level sufficient",
        "No visible fluid leaks",
        "Air filters clean",
        "Moisture drained from system",
        "Hoses and fittings secure",
        "Guards and covers secure",
        "Gauges operating properly",
        "Emergency shutdown functional"
      ]
    },
    {
      title: "Dust Collector – Daily Inspection Checklist",
      items: [
        "Guards and access panels secure",
        "Emergency stops functional",
        "Warning labels legible",
        "No visible damage or leaks",
        "Filter bags/cartridges intact",
        "Differential pressure normal",
        "Hopper free of buildup",
        "Dust discharge operating",
        "Control panel indicators normal",
        "No alarm conditions present"
      ]
    },
    {
      title: "Blast Machine – Daily Inspection Checklist",
      items: [
        "Machine frame and guards intact",
        "Emergency stop functional",
        "Access doors secured",
        "Screens free of blockage",
        "Magnetic separator clean",
        "Conveyors operating smoothly",
        "Air lines free of leaks",
        "Bearings lubricated",
        "No abnormal vibration or noise"
      ]
    },
    {
      title: "Vacuum – Daily Inspection Checklist",
      items: [
        "Blower operating normally (28\" Hg)",
        "Hoses free of damage",
        "Boom and joints operate smoothly",
        "Tank free of excessive buildup",
        "Rear door seals intact",
        "Door latching secure",
        "Sludge pump functional",
        "Valves operate smoothly",
        "No hydraulic leaks observed"
      ]
    }
  ]
};

const state = {
  project: "",
  inspectorName: "",
  inspectorDate: todayISO(),
  signature: null, // dataURL (png)
  pages: CONFIG.pages.map(p => ({
    na: false,
    additionalComments: "",
    items: p.items.map(() => ({ status: "", comment: "" })),
    // 4 slots max per page; each slot: { itemIndex, dataURL, bytes, mime }
    photoSlots: [null, null, null, null]
  }))
};

const els = {
  project: document.getElementById("projectName"),
  inspectorName: document.getElementById("inspectorName"),
  inspectorDate: document.getElementById("inspectorDate"),
  sigPreview: document.getElementById("sigPreview"),
  sigBtn: document.getElementById("sigBtn"),
  pages: document.getElementById("pages"),
  sigModal: document.getElementById("sigModal"),
  sigCanvas: document.getElementById("sigCanvas"),
  sigTitle: document.getElementById("sigTitle"),
  sigClose: document.getElementById("sigClose"),
  sigClear: document.getElementById("sigClear"),
  sigCancel: document.getElementById("sigCancel"),
  sigSave: document.getElementById("sigSave"),
};

let attachCtx = null; // { pageIndex, itemIndex }
let sigDrawing = false;
let sigLast = null;

function todayISO() {
  const d = new Date();
  const y = d.getFullYear();
  const m = String(d.getMonth()+1).padStart(2,"0");
  const day = String(d.getDate()).padStart(2,"0");
  return `${y}-${m}-${day}`;
}

function isoToMMDDYYYY(iso) {
  if (!iso) return "";
  const parts = iso.split("-");
  if (parts.length !== 3) return iso;
  return `${parts[1]}/${parts[2]}/${parts[0]}`;
}

function escapeHtml(s) {
  return (s||"").replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}

function setDisabled(disabled) {
  document.querySelectorAll('button, input, select, textarea').forEach(el => {
    if (el.closest('.modal')) return; // keep modal usable
    el.disabled = disabled;
  });
}

function wireTopControls() {
  document.body.addEventListener("click", (e) => {
    const btn = e.target.closest("[data-action]");
    if (!btn) return;
    const action = btn.getAttribute("data-action");
    if (action === "reset") doReset();
    if (action === "save") doSave();
  });

  els.project.addEventListener("change", () => state.project = els.project.value);
  els.inspectorName.addEventListener("input", () => state.inspectorName = els.inspectorName.value);
  els.inspectorDate.addEventListener("change", () => state.inspectorDate = els.inspectorDate.value);

  els.sigBtn.addEventListener("click", () => openSignatureModal());
}

function renderProjectOptions() {
  els.project.innerHTML = PROJECT_OPTIONS.map(p => `<option value="${escapeHtml(p)}">${escapeHtml(p)}</option>`).join("");
  els.project.value = state.project || PROJECT_OPTIONS[0];
}

function renderPages() {
  const html = CONFIG.pages.map((p, pageIndex) => {
    const pState = state.pages[pageIndex];
    const naChecked = pState.na ? "checked" : "";
    const itemsHtml = p.items.map((label, itemIndex) => {
      const iState = pState.items[itemIndex];
      const okOn = iState.status === "ok" ? "active" : "";
      const needsOn = iState.status === "needs" ? "active" : "";
      const commentVal = escapeHtml(iState.comment);

      const slot = findSlotByItem(pageIndex, itemIndex);
      const thumb = slot ? `<img class="thumb" src="${slot.dataURL}" alt="photo thumbnail"/>` : "";
      const removeBtn = slot ? `<button class="btn tiny danger" data-remove-photo="${pageIndex}:${itemIndex}" type="button">Remove photo</button>` : "";

      return `
        <div class="item-row" data-page="${pageIndex}" data-item="${itemIndex}">
          <div class="item-label">${escapeHtml(label)}</div>

          <div class="item-controls">
            <div class="seg">
              <button class="seg-btn ${okOn}" type="button" data-set-status="${pageIndex}:${itemIndex}:ok">OK</button>
              <button class="seg-btn ${needsOn}" type="button" data-set-status="${pageIndex}:${itemIndex}:needs">Needs attention</button>
            </div>

            <div class="photo-controls">
              <button class="btn tiny" type="button" data-attach-photo="${pageIndex}:${itemIndex}">Attach photo</button>
              ${removeBtn}
              ${thumb}
            </div>
          </div>

          <textarea class="textarea" rows="2" placeholder="Comments" data-comment="${pageIndex}:${itemIndex}">${commentVal}</textarea>
        </div>
      `;
    }).join("");

    return `
      <section class="card">
        <div class="card-title">
          <div>${escapeHtml(p.title)}</div>
          <label class="na-toggle">
            <input type="checkbox" data-na="${pageIndex}" ${naChecked}/>
            <span>N/A</span>
          </label>
        </div>

        <div class="hint">Up to 4 photos per checklist page (matches the 4 photo boxes on the PDF).</div>

        <div class="items ${pState.na ? "hidden" : ""}">
          ${itemsHtml}
        </div>

        <div class="row">
          <label class="label" for="addC_${pageIndex}">Additional comments</label>
          <textarea id="addC_${pageIndex}" class="textarea" rows="3" data-additional="${pageIndex}" placeholder="Additional comments for this checklist page">${escapeHtml(pState.additionalComments)}</textarea>
        </div>
      </section>
    `;
  }).join("");

  els.pages.innerHTML = html;

}

function handlePagesClick(e) {
  const btnStatus = e.target.closest("[data-set-status]");
  if (btnStatus) {
    const [pageIndexS,itemIndexS,status] = btnStatus.getAttribute("data-set-status").split(":");
    const pageIndex = Number(pageIndexS), itemIndex = Number(itemIndexS);
    if (state.pages[pageIndex].na) return;
    state.pages[pageIndex].items[itemIndex].status = status;
    rerenderPageItems(pageIndex);
    return;
  }

  const btnAttach = e.target.closest("[data-attach-photo]");
  if (btnAttach) {
    const [pageIndexS,itemIndexS] = btnAttach.getAttribute("data-attach-photo").split(":");
    const pageIndex = Number(pageIndexS), itemIndex = Number(itemIndexS);
    if (state.pages[pageIndex].na) return;
    openPhotoPicker(pageIndex, itemIndex);
    return;
  }

  const btnRemove = e.target.closest("[data-remove-photo]");
  if (btnRemove) {
    const [pageIndexS,itemIndexS] = btnRemove.getAttribute("data-remove-photo").split(":");
    removePhoto(Number(pageIndexS), Number(itemIndexS));
    rerenderPageItems(Number(pageIndexS));
    return;
  }
}

function handlePagesInput(e) {
  const ta = e.target.closest("[data-comment]");
  if (ta) {
    const [pageIndexS,itemIndexS] = ta.getAttribute("data-comment").split(":");
    const pageIndex = Number(pageIndexS), itemIndex = Number(itemIndexS);
    state.pages[pageIndex].items[itemIndex].comment = ta.value;
    return;
  }

  const add = e.target.closest("[data-additional]");
  if (add) {
    const pageIndex = Number(add.getAttribute("data-additional"));
    state.pages[pageIndex].additionalComments = add.value;
  }
}

function handlePagesChange(e) {
  const na = e.target.closest("[data-na]");
  if (na) {
    const pageIndex = Number(na.getAttribute("data-na"));
    state.pages[pageIndex].na = na.checked;

    if (na.checked) {
      // Clear statuses + photos to avoid accidental carryover
      state.pages[pageIndex].items.forEach(i => { i.status = ""; i.comment = ""; });
      state.pages[pageIndex].photoSlots = [null,null,null,null];
    }

    rerenderWholePages();
  }
}

function rerenderWholePages() {
  // Remove old listeners by re-rendering fully
  renderPages();
}

function rerenderPageItems(pageIndex) {
  // Simple approach: full re-render (small form, 4 pages)
  rerenderWholePages();
}

function openPhotoPicker(pageIndex, itemIndex) {
  attachCtx = { pageIndex, itemIndex };
  let input = document.getElementById("photoFileInput");
  if (!input) {
    input = document.createElement("input");
    input.type = "file";
    input.accept = "image/*";
    input.id = "photoFileInput";
    input.style.display = "none";
    document.body.appendChild(input);
    input.addEventListener("change", async () => {
      const file = input.files && input.files[0];
      input.value = "";
      if (!file || !attachCtx) return;

      const { pageIndex, itemIndex } = attachCtx;
      attachCtx = null;

      try {
        const dataURL = await fileToDataURL(file);
        const bytes = dataURLToBytes(dataURL);
        const mime = (dataURL.match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,/)||[])[1] || "image/png";
        addOrReplacePhoto(pageIndex, itemIndex, { dataURL, bytes, mime });
        rerenderPageItems(pageIndex);
      } catch (err) {
        alert("Could not read the image file.");
        console.error(err);
      }
    });
  }
  input.click();
}

function findSlotByItem(pageIndex, itemIndex) {
  const slots = state.pages[pageIndex].photoSlots;
  for (let s=0;s<slots.length;s++) {
    if (slots[s] && slots[s].itemIndex === itemIndex) return slots[s];
  }
  return null;
}

function addOrReplacePhoto(pageIndex, itemIndex, photoObj) {
  const pState = state.pages[pageIndex];

  // Replace if already exists
  for (let s=0;s<pState.photoSlots.length;s++) {
    if (pState.photoSlots[s] && pState.photoSlots[s].itemIndex === itemIndex) {
      pState.photoSlots[s] = { itemIndex, ...photoObj };
      return;
    }
  }

  // Find free slot
  const free = pState.photoSlots.findIndex(x => !x);
  if (free === -1) {
    alert("This checklist page already has 4 photos. Remove one to attach a new photo.");
    return;
  }
  pState.photoSlots[free] = { itemIndex, ...photoObj };
}

function removePhoto(pageIndex, itemIndex) {
  const pState = state.pages[pageIndex];
  for (let s=0;s<pState.photoSlots.length;s++) {
    if (pState.photoSlots[s] && pState.photoSlots[s].itemIndex === itemIndex) {
      pState.photoSlots[s] = null;
      return;
    }
  }
}

function fileToDataURL(file) {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(r.result);
    r.onerror = reject;
    r.readAsDataURL(file);
  });
}

function dataURLToBytes(dataURL) {
  const base64 = dataURL.split(",")[1] || "";
  const bin = atob(base64);
  const len = bin.length;
  const bytes = new Uint8Array(len);
  for (let i=0;i<len;i++) bytes[i] = bin.charCodeAt(i);
  return bytes;
}

function openSignatureModal() {
  els.sigTitle.textContent = "Signature (applies to all pages)";
  els.sigModal.classList.remove("hidden");
  setupSigCanvas();
}

function closeSignatureModal() {
  els.sigModal.classList.add("hidden");
}

function setupSigCanvas() {
  const canvas = els.sigCanvas;
  const rect = canvas.getBoundingClientRect();

  const dpr = window.devicePixelRatio || 1;
  canvas.width = Math.floor(rect.width * dpr);
  canvas.height = Math.floor(180 * dpr);
  canvas.style.height = "180px";

  const ctx = canvas.getContext("2d");
  // reset transform, then scale for DPR
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  ctx.lineWidth = 2;
  ctx.lineCap = "round";
  ctx.strokeStyle = "#111";

  // white background
  ctx.fillStyle = "#fff";
  ctx.fillRect(0, 0, rect.width, 180);

  // if existing signature, draw it
  if (state.signature) {
    const img = new Image();
    img.onload = () => {
      const r = canvas.getBoundingClientRect();
      const ctx2 = canvas.getContext("2d");
      ctx2.setTransform(dpr, 0, 0, dpr, 0, 0);
      ctx2.drawImage(img, 0, 0, r.width, 180);
    };
    img.src = state.signature;
  }
}

function sigGetPos(ev) {
  const r = els.sigCanvas.getBoundingClientRect();
  return { x: ev.clientX - r.left, y: ev.clientY - r.top };
}

function onSigPointerDown(ev) {
  ev.preventDefault();
  sigDrawing = true;
  sigLast = sigGetPos(ev);
  try { els.sigCanvas.setPointerCapture(ev.pointerId); } catch (_) {}
}

function onSigPointerMove(ev) {
  if (!sigDrawing) return;
  ev.preventDefault();
  const pos = sigGetPos(ev);
  const ctx = els.sigCanvas.getContext("2d");
  ctx.beginPath();
  ctx.moveTo(sigLast.x, sigLast.y);
  ctx.lineTo(pos.x, pos.y);
  ctx.stroke();
  sigLast = pos;
}

function onSigPointerUp(ev) {
  if (!sigDrawing) return;
  ev.preventDefault();
  sigDrawing = false;
  sigLast = null;
}

function updateSigPreview() {() {
  if (state.signature) {
    els.sigPreview.innerHTML = `<img src="${state.signature}" alt="signature preview" />`;
  } else {
    els.sigPreview.textContent = "No signature saved";
  }
}

function clearSignatureCanvas() {
  const canvas = els.sigCanvas;
  const ctx = canvas.getContext("2d");
  const rect = canvas.getBoundingClientRect();
  ctx.fillStyle = "#fff";
  ctx.fillRect(0,0,rect.width,180);
}

function saveSignatureFromCanvas() {
  const canvas = els.sigCanvas;
  // export at screen size, not internal size
  const tmp = document.createElement("canvas");
  const r = canvas.getBoundingClientRect();
  tmp.width = Math.floor(r.width);
  tmp.height = Math.floor(180);
  const tctx = tmp.getContext("2d");
  tctx.fillStyle = "#fff";
  tctx.fillRect(0,0,tmp.width,tmp.height);
  tctx.drawImage(canvas, 0, 0, tmp.width, tmp.height);

  state.signature = tmp.toDataURL("image/png");
  updateSigPreview();
  closeSignatureModal();
}

function wrapLineToWidth(text, maxWidth, font, size) {
  // Returns a single line truncated with ellipsis to fit width
  if (!text) return "";
  let t = String(text).replace(/\s+/g, " ").trim();
  if (!t) return "";
  const w = font.widthOfTextAtSize(t, size);
  if (w <= maxWidth) return t;

  const ell = "…";
  let lo = 0, hi = t.length;
  while (lo < hi) {
    const mid = Math.ceil((lo + hi) / 2);
    const sub = t.slice(0, mid) + ell;
    if (font.widthOfTextAtSize(sub, size) <= maxWidth) lo = mid;
    else hi = mid - 1;
  }
  return t.slice(0, lo) + ell;
}

function fitSizeToWidth(text, maxWidth, font, startSize, minSize) {
  let size = startSize;
  while (size > minSize) {
    if (font.widthOfTextAtSize(text, size) <= maxWidth) return size;
    size -= 0.5;
  }
  return minSize;
}

function drawImageContain(page, image, box) {
  const pad = 2;
  const bx = box.x + pad;
  const by = box.y + pad;
  const bw = box.w - pad*2;
  const bh = box.h - pad*2;

  const iw = image.width;
  const ih = image.height;

  const scale = Math.min(bw/iw, bh/ih);
  const w = iw * scale;
  const h = ih * scale;

  const x = bx + (bw - w) / 2;
  const y = by + (bh - h) / 2;

  page.drawImage(image, { x, y, width: w, height: h });
}

async function doSave() {
  setDisabled(true);
  try {
    const templateBytes = await fetch(CONFIG.templateFile).then(r => r.arrayBuffer());
    const { PDFDocument, rgb, StandardFonts } = PDFLib;

    const pdfDoc = await PDFDocument.load(templateBytes);
    const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
    const fontBold = await pdfDoc.embedFont(StandardFonts.HelveticaBold);

    const pages = pdfDoc.getPages();
    const { width, height } = pages[0].getSize();

    const stampDateISO = state.inspectorDate || todayISO();
    const stampDate = isoToMMDDYYYY(stampDateISO);
    const naText = `Not applicable for today ${stampDate}`;

    // Pre-embed signature once (if any)
    let sigImg = null;
    if (state.signature) {
      const sigBytes = dataURLToBytes(state.signature);
      sigImg = await pdfDoc.embedPng(sigBytes);
    }

    // Fill pages
    for (let pageIndex=0; pageIndex<CONFIG.pages.length; pageIndex++) {
      const page = pages[pageIndex];
      const pConf = CONFIG.pages[pageIndex];
      const pState = state.pages[pageIndex];

      // Project name (top)
      if (state.project && state.project.trim()) {
        const t = state.project.trim();
        const baseSize = CONFIG.fields.project.size;
        const size = fitSizeToWidth(t, CONFIG.fields.project.w, fontBold, baseSize, 8);
        page.drawText(t, {
          x: CONFIG.fields.project.x,
          y: CONFIG.fields.project.y,
          size,
          font: fontBold,
          color: rgb(0,0,0)
        });
      }

      // Table rows
      const rowCenters = CONFIG.table.rowCentersTop;
      for (let r=0; r<pConf.items.length; r++) {
        const y = height - rowCenters[r];
        let status = pState.items[r]?.status || "";
        let comment = pState.items[r]?.comment || "";

        if (pState.na) {
          status = "";
          comment = naText;
        }

        if (status === "ok") {
          page.drawText("X", { x: CONFIG.table.okCenterX - 4, y: y - 5, size: 12, font: fontBold, color: rgb(0,0,0) });
        } else if (status === "needs") {
          page.drawText("X", { x: CONFIG.table.needsCenterX - 4, y: y - 5, size: 12, font: fontBold, color: rgb(0,0,0) });
        }

        if (comment && comment.trim()) {
          const line = wrapLineToWidth(comment, CONFIG.table.commentW, font, 9);
          page.drawText(line, { x: CONFIG.table.commentX, y: y - 4, size: 9, font, color: rgb(0,0,0) });
        }
      }

      // Additional comments (bottom)
      if (pState.additionalComments && pState.additionalComments.trim()) {
        const lines = String(pState.additionalComments).split(/\r?\n/).map(s => s.trim()).filter(Boolean);
        // flatten into words and wrap
        const joined = lines.join(" ");
        const maxW = CONFIG.fields.additionalComments.w;
        const size = CONFIG.fields.additionalComments.size;
        const lineH = CONFIG.fields.additionalComments.lineHeight;
        const maxLines = CONFIG.fields.additionalComments.maxLines;

        const wrapped = wrapParagraph(joined, maxW, font, size, maxLines);
        const yTop = CONFIG.fields.additionalComments.yTop;
        for (let i=0;i<wrapped.length;i++) {
          page.drawText(wrapped[i], { x: CONFIG.fields.additionalComments.x, y: yTop - i*lineH, size, font, color: rgb(0,0,0) });
        }
      }

      // Name + Date
      if (state.inspectorName && state.inspectorName.trim()) {
        const name = state.inspectorName.trim();
        const size = fitSizeToWidth(name, CONFIG.fields.name.w, font, CONFIG.fields.name.size, 8);
        page.drawText(name, { x: CONFIG.fields.name.x, y: CONFIG.fields.name.y, size, font, color: rgb(0,0,0) });
      }
      if (stampDate) {
        page.drawText(stampDate, { x: CONFIG.fields.date.x, y: CONFIG.fields.date.y, size: CONFIG.fields.date.size, font, color: rgb(0,0,0) });
      }

      // Signature (image)
      if (sigImg) {
        drawImageContain(page, sigImg, CONFIG.fields.signature);
      }

      // Photos (4 slots)
      if (!pState.na) {
        for (let s=0; s<CONFIG.photos.length; s++) {
          const slot = pState.photoSlots[s];
          if (!slot) continue;

          const itemLabel = pConf.items[slot.itemIndex] || "";
          const hdr = CONFIG.photos[s].header;
          const box = CONFIG.photos[s].box;

          // Header: inspection item title (blank if no photo)
          if (itemLabel) {
            const baseSize = 9;
            const trimmed = String(itemLabel).trim();
            const size = fitSizeToWidth(trimmed, hdr.w - 6, fontBold, baseSize, 6);
            const line = wrapLineToWidth(trimmed, hdr.w - 6, fontBold, size);
            page.drawText(line, { x: hdr.x + 3, y: hdr.y + 4, size, font: fontBold, color: rgb(0,0,0) });
          }

          // Image into box
          let img;
          if (slot.mime.includes("jpeg") || slot.mime.includes("jpg")) {
            img = await pdfDoc.embedJpg(slot.bytes);
          } else {
            img = await pdfDoc.embedPng(slot.bytes);
          }
          drawImageContain(page, img, box);
        }
      } else {
        // If N/A, ensure photo headers stay blank by doing nothing.
      }
    }

    // Create a fresh output PDF by copying pages (drops any catalog-level JS/name trees)
    const outDoc = await PDFDocument.create();
    const copied = await outDoc.copyPages(pdfDoc, pdfDoc.getPageIndices());
    copied.forEach(p => outDoc.addPage(p));
    const outBytes = await outDoc.save();

    downloadBytes(outBytes, `JAGD_Maintenance_Checklists_${stampDate.replaceAll("/","-")}.pdf`);
  } catch (err) {
    console.error(err);
    alert("Could not generate the PDF. Check console for details.");
  } finally {
    setDisabled(false);
  }
}

function wrapParagraph(text, maxWidth, font, size, maxLines) {
  const words = String(text || "").split(/\s+/).filter(Boolean);
  const lines = [];
  let line = "";
  for (const w of words) {
    const test = line ? (line + " " + w) : w;
    if (font.widthOfTextAtSize(test, size) <= maxWidth) {
      line = test;
      continue;
    }
    if (line) lines.push(line);
    line = w;
    if (lines.length >= maxLines) break;
  }
  if (lines.length < maxLines && line) lines.push(line);

  // If we overflowed, add ellipsis to last line
  if (lines.length === maxLines && words.length) {
    lines[maxLines-1] = wrapLineToWidth(lines[maxLines-1], maxWidth, font, size);
  }
  return lines;
}

function downloadBytes(bytes, filename) {
  const blob = new Blob([bytes], { type: "application/pdf" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function doReset() {
  if (!confirm("Reset all fields?")) return;

  state.project = PROJECT_OPTIONS[0];
  state.inspectorName = "";
  state.inspectorDate = todayISO();
  state.signature = null;

  state.pages = CONFIG.pages.map(p => ({
    na: false,
    additionalComments: "",
    items: p.items.map(() => ({ status: "", comment: "" })),
    photoSlots: [null,null,null,null]
  }));

  // Reset UI inputs
  els.project.value = state.project;
  els.inspectorName.value = "";
  els.inspectorDate.value = state.inspectorDate;
  updateSigPreview();

  rerenderWholePages();
}

function wireSignatureModal() {
  els.sigClose.addEventListener("click", closeSignatureModal);
  els.sigCancel.addEventListener("click", closeSignatureModal);
  els.sigClear.addEventListener("click", () => clearSignatureCanvas());
  els.sigSave.addEventListener("click", () => saveSignatureFromCanvas());

  // Signature canvas pointer handlers (wire once)
  els.sigCanvas.addEventListener("pointerdown", onSigPointerDown);
  els.sigCanvas.addEventListener("pointermove", onSigPointerMove);
  els.sigCanvas.addEventListener("pointerup", onSigPointerUp);
  els.sigCanvas.addEventListener("pointercancel", onSigPointerUp);


  // click outside
  els.sigModal.addEventListener("click", (e) => {
    if (e.target === els.sigModal) closeSignatureModal();
  });
}

function init() {
  wireTopControls();
  wireSignatureModal();

  // Delegated handlers (wire once)
  els.pages.addEventListener("click", handlePagesClick);
  els.pages.addEventListener("input", handlePagesInput);
  els.pages.addEventListener("change", handlePagesChange);

  renderProjectOptions();
  els.inspectorDate.value = state.inspectorDate;

  updateSigPreview();
  rerenderWholePages();
}

init();
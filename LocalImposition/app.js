(function () {
  const state = {
    busy: false,
    lastResult: null,
    inboxFiles: []
  };

  const elements = {
    form: document.getElementById("imposition-form"),
    fileInput: document.getElementById("source-pdf"),
    dropZone: document.getElementById("drop-zone"),
    uploadPanel: document.getElementById("upload-panel"),
    inboxPanel: document.getElementById("inbox-panel"),
    inboxFile: document.getElementById("inbox-file"),
    refreshInbox: document.getElementById("refresh-inbox"),
    inboxSummary: document.getElementById("inbox-summary"),
    selectedFile: document.getElementById("selected-file"),
    generateButton: document.getElementById("generate-button"),
    status: document.getElementById("status"),
    metricSource: document.getElementById("metric-source"),
    metricSheet: document.getElementById("metric-sheet"),
    metricLayout: document.getElementById("metric-layout"),
    metricOrdering: document.getElementById("metric-ordering"),
    metricOutput: document.getElementById("metric-output"),
    metricAdjustments: document.getElementById("metric-adjustments"),
    metricNote: document.getElementById("metric-note"),
    cutWidth: document.getElementById("cut-width"),
    cutHeight: document.getElementById("cut-height"),
    bleedWidth: document.getElementById("bleed-width"),
    bleedHeight: document.getElementById("bleed-height"),
    sheetWidth: document.getElementById("sheet-width"),
    sheetHeight: document.getElementById("sheet-height"),
    gapHorizontal: document.getElementById("gap-horizontal"),
    gapVertical: document.getElementById("gap-vertical"),
    sheetOrientation: document.getElementById("sheet-orientation"),
    bestFit: document.getElementById("best-fit"),
    duplex: document.getElementById("duplex"),
    autoCorrect: document.getElementById("auto-correct"),
    artRotation: document.getElementById("art-rotation"),
    bindingEdge: document.getElementById("binding-edge"),
    rotateFirst: document.getElementById("rotate-first"),
    shiftX: document.getElementById("shift-x"),
    shiftY: document.getElementById("shift-y"),
    sourceModeInputs: Array.from(document.querySelectorAll("input[name='sourceMode']"))
  };

  bindEvents();
  updateSourcePanels();
  updateSummary();
  loadInboxFiles(true);

  function bindEvents() {
    elements.form.addEventListener("submit", handleGenerate);
    elements.fileInput.addEventListener("change", handleFileChange);
    elements.refreshInbox.addEventListener("click", () => loadInboxFiles(false));
    elements.inboxFile.addEventListener("change", handleInboxChange);

    for (const input of elements.sourceModeInputs) {
      input.addEventListener("change", () => {
        state.lastResult = null;
        updateSourcePanels();
        updateSummary();
      });
    }

    for (const control of elements.form.querySelectorAll("input, select")) {
      if (control === elements.fileInput || elements.sourceModeInputs.includes(control)) continue;
      control.addEventListener("input", updateSummary);
      control.addEventListener("change", updateSummary);
    }

    ["dragenter", "dragover"].forEach((eventName) => {
      elements.dropZone.addEventListener(eventName, (event) => {
        event.preventDefault();
        elements.dropZone.classList.add("is-dragover");
      });
    });

    ["dragleave", "dragend", "drop"].forEach((eventName) => {
      elements.dropZone.addEventListener(eventName, (event) => {
        event.preventDefault();
        elements.dropZone.classList.remove("is-dragover");
      });
    });

    elements.dropZone.addEventListener("drop", (event) => {
      const [file] = Array.from(event.dataTransfer?.files || []);
      if (!file) return;
      if (file.type !== "application/pdf" && !file.name.toLowerCase().endsWith(".pdf")) {
        setStatus("Only PDF files are supported.", "error");
        return;
      }

      const transfer = new DataTransfer();
      transfer.items.add(file);
      elements.fileInput.files = transfer.files;
      handleFileChange();
    });
  }

  function handleFileChange() {
    const file = elements.fileInput.files?.[0] || null;
    state.lastResult = null;

    if (!file) {
      elements.selectedFile.textContent = "No upload selected.";
      if (getSourceMode() === "upload") {
        setStatus("Choose a PDF to upload, or switch back to the inbox folder source.", "info");
      }
      updateSummary();
      return;
    }

    elements.selectedFile.textContent = `${file.name} • ${formatFileSize(file.size)}`;
    if (getSourceMode() === "upload") {
      setStatus(`Ready to upload ${file.name} to the local backend.`, "success");
    }
    updateSummary();
  }

  function handleInboxChange() {
    state.lastResult = null;
    const entry = getSelectedInboxEntry();
    elements.inboxSummary.textContent = entry
      ? `${entry.name} • ${formatFileSize(entry.size)}`
      : "No inbox PDF selected.";
    if (getSourceMode() === "inbox" && entry) {
      setStatus(`Ready to use ${entry.name} from LocalImposition/inbox.`, "success");
    }
    updateSummary();
  }

  async function handleGenerate(event) {
    event.preventDefault();
    if (state.busy) return;

    const sourceMode = getSourceMode();
    const uploadFile = elements.fileInput.files?.[0] || null;
    const inboxEntry = getSelectedInboxEntry();

    if (sourceMode === "upload" && !uploadFile) {
      setStatus("Select a PDF before generating the imposed output.", "error");
      return;
    }

    if (sourceMode === "inbox" && !inboxEntry) {
      setStatus("Choose a PDF from LocalImposition/inbox before generating the imposed output.", "error");
      return;
    }

    try {
      setBusy(true);
      if (sourceMode === "upload") {
        setStatus(
          uploadFile.size > 1024 * 1024 * 1024
            ? "Uploading a large PDF to the local backend. Keep this tab open while the server processes it."
            : "Uploading the PDF to the local backend...",
          "info"
        );
      } else {
        setStatus(`Processing ${inboxEntry.name} directly from LocalImposition/inbox...`, "info");
      }

      const config = readConfig();
      const formData = new FormData();
      for (const [key, value] of Object.entries(config)) {
        formData.set(key, String(value));
      }

      if (sourceMode === "upload") {
        formData.set("sourcePdf", uploadFile, uploadFile.name);
      } else {
        formData.set("inboxFile", inboxEntry.name);
      }

      const response = await fetch("/api/impose", {
        method: "POST",
        body: formData
      });

      const payload = await response.json().catch(() => null);
      if (!response.ok || !payload) {
        throw new Error(payload?.error || `Imposition request failed with status ${response.status}.`);
      }

      state.lastResult = payload;
      setStatus(
        `Imposition finished: ${payload.outputSheetCount} sheet(s) at ${payload.cols} x ${payload.rows} up. Starting the download.`,
        "success"
      );
      updateSummary();
      window.location.href = payload.downloadUrl;
    } catch (error) {
      const isNetworkError = error instanceof TypeError;
      setStatus(
        isNetworkError
          ? "The local backend is not reachable. Start `server.py` and reload this page."
          : error.message || "Failed to impose the PDF.",
        "error"
      );
    } finally {
      setBusy(false);
    }
  }

  async function loadInboxFiles(quiet) {
    try {
      const previousSelection = elements.inboxFile.value;
      const response = await fetch("/api/files");
      const payload = await response.json().catch(() => null);
      if (!response.ok || !payload) {
        throw new Error(payload?.error || `Inbox request failed with status ${response.status}.`);
      }

      state.inboxFiles = Array.isArray(payload.files) ? payload.files : [];
      renderInboxOptions(previousSelection);

      if (!quiet && getSourceMode() === "inbox") {
        setStatus(
          state.inboxFiles.length
            ? `Found ${state.inboxFiles.length} PDF file(s) in LocalImposition/inbox.`
            : "No PDFs are in LocalImposition/inbox yet.",
          "info"
        );
      }
    } catch (error) {
      state.inboxFiles = [];
      renderInboxOptions("");
      if (!quiet) {
        const isNetworkError = error instanceof TypeError;
        setStatus(
          isNetworkError
            ? "Could not reach the local backend to read LocalImposition/inbox. Start `server.py` and reload this page."
            : error.message || "Could not read the inbox directory.",
          "error"
        );
      }
    }

    updateSummary();
  }

  function renderInboxOptions(preferredName) {
    const files = state.inboxFiles;
    elements.inboxFile.innerHTML = "";

    if (!files.length) {
      const option = document.createElement("option");
      option.value = "";
      option.textContent = "No PDFs found in LocalImposition/inbox";
      elements.inboxFile.append(option);
      elements.inboxFile.disabled = true;
      elements.inboxSummary.textContent = "Place a PDF in LocalImposition/inbox and press Refresh list.";
      return;
    }

    for (const entry of files) {
      const option = document.createElement("option");
      option.value = entry.name;
      option.textContent = `${entry.name} • ${formatFileSize(entry.size)}`;
      elements.inboxFile.append(option);
    }

    const selectedName = files.some((entry) => entry.name === preferredName)
      ? preferredName
      : files[0].name;
    elements.inboxFile.value = selectedName;
    elements.inboxFile.disabled = false;

    const selectedEntry = getSelectedInboxEntry();
    elements.inboxSummary.textContent = selectedEntry
      ? `${selectedEntry.name} • ${formatFileSize(selectedEntry.size)}`
      : "No inbox PDF selected.";
  }

  function updateSourcePanels() {
    const sourceMode = getSourceMode();
    elements.inboxPanel.hidden = sourceMode !== "inbox";
    elements.uploadPanel.hidden = sourceMode !== "upload";
  }

  function getSourceMode() {
    return elements.sourceModeInputs.find((input) => input.checked)?.value || "inbox";
  }

  function getSelectedInboxEntry() {
    const name = elements.inboxFile.value;
    return state.inboxFiles.find((entry) => entry.name === name) || null;
  }

  function readConfig() {
    const cutWidthIn = readPositiveNumber(elements.cutWidth, "Cut width");
    const cutHeightIn = readPositiveNumber(elements.cutHeight, "Cut height");
    const bleedWidthIn = Math.max(cutWidthIn, readPositiveNumber(elements.bleedWidth, "Bleed width"));
    const bleedHeightIn = Math.max(cutHeightIn, readPositiveNumber(elements.bleedHeight, "Bleed height"));
    const sheetWidthIn = readPositiveNumber(elements.sheetWidth, "Sheet width");
    const sheetHeightIn = readPositiveNumber(elements.sheetHeight, "Sheet height");
    const gapHorizontalIn = readNonNegativeNumber(elements.gapHorizontal, "Horizontal gap");
    const gapVerticalIn = readNonNegativeNumber(elements.gapVertical, "Vertical gap");
    const mode = String(elements.form.elements.mode.value || "repeat");

    return {
      mode: mode === "cutAndStack" ? "cutAndStack" : "repeat",
      cutWidthIn,
      cutHeightIn,
      bleedWidthIn,
      bleedHeightIn,
      sheetWidthIn,
      sheetHeightIn,
      gapHorizontalIn,
      gapVerticalIn,
      sheetOrientation: String(elements.sheetOrientation.value || "auto"),
      bestFit: elements.bestFit.checked,
      duplex: elements.duplex.checked,
      autoCorrectArtOrientation: elements.autoCorrect.checked,
      artRotation: String(elements.artRotation.value || "None"),
      bindingEdge: String(elements.bindingEdge.value || "Left"),
      rotateFirstColumnOrRow: elements.rotateFirst.checked,
      imageShiftXIn: readSignedNumber(elements.shiftX),
      imageShiftYIn: readSignedNumber(elements.shiftY)
    };
  }

  function updateSummary() {
    try {
      const config = readConfig();
      const resolvedSheet = resolveSheetConfiguration(config);
      const layout =
        resolvedSheet.capacity > 0
          ? planLayout(
              resolvedSheet.sheetWIn,
              resolvedSheet.sheetHIn,
              config.cutWidthIn,
              config.cutHeightIn,
              config.gapHorizontalIn,
              config.gapVerticalIn,
              resolvedSheet.capacity
            )
          : null;
      const sourceMode = getSourceMode();
      const uploadFile = elements.fileInput.files?.[0] || null;
      const inboxEntry = getSelectedInboxEntry();
      const sourceLabel =
        sourceMode === "inbox"
          ? inboxEntry
            ? `${inboxEntry.name} • ${formatFileSize(inboxEntry.size)}`
            : "No inbox PDF selected."
          : uploadFile
            ? `${uploadFile.name} • ${formatFileSize(uploadFile.size)}`
            : "No upload selected.";

      elements.metricSource.textContent = sourceLabel;
      elements.metricSheet.textContent = `${formatNumber(resolvedSheet.sheetWIn)} x ${formatNumber(resolvedSheet.sheetHIn)} in (${resolvedSheet.actualOrientation})`;
      elements.metricLayout.textContent = layout
        ? `${layout.cols} x ${layout.rows} up (${layout.cols * layout.rows} positions)`
        : "Current sizes do not fit on the selected sheet.";
      elements.metricOrdering.textContent =
        config.mode === "cutAndStack"
          ? "Position-major cut-and-stack ordering across imposed sheets."
          : "Each source page repeats into every position on its own sheet.";
      elements.metricOutput.textContent = state.lastResult
        ? `${state.lastResult.outputSheetCount} imposed sheet(s) from ${state.lastResult.pageCount} source page(s)`
        : sourceMode === "inbox"
          ? "Source page count is determined by the backend after it opens the selected inbox file."
          : uploadFile
            ? "Source page count is determined by the backend during processing."
            : "Choose a source PDF to generate output.";
      elements.metricAdjustments.textContent = buildAdjustmentSummary(config, resolvedSheet);
      elements.metricNote.textContent = buildNoteSummary(config, layout, resolvedSheet, sourceMode, sourceLabel, state.lastResult);

      if (!state.busy && (!elements.status.dataset.state || elements.status.dataset.state === "info")) {
        setStatus(
          layout
            ? "Layout is valid. Generate the imposed PDF when you are ready."
            : "Adjust the sheet, cut, bleed, or gap values until the layout fits.",
          "info"
        );
      }
    } catch (error) {
      elements.metricSource.textContent = getSourceMode() === "inbox"
        ? getSelectedInboxEntry()
          ? `${getSelectedInboxEntry().name} • ${formatFileSize(getSelectedInboxEntry().size)}`
          : "No inbox PDF selected."
        : elements.fileInput.files?.[0]
          ? `${elements.fileInput.files[0].name} • ${formatFileSize(elements.fileInput.files[0].size)}`
          : "No upload selected.";
      elements.metricSheet.textContent = "Waiting for valid dimensions.";
      elements.metricLayout.textContent = "Waiting for valid dimensions.";
      elements.metricOrdering.textContent = "Choose repeat or cut and stack.";
      elements.metricOutput.textContent = "Choose a source PDF to generate output.";
      elements.metricAdjustments.textContent = "No extra adjustments enabled.";
      elements.metricNote.textContent = error.message || "Enter positive dimensions to continue.";
    }
  }

  function resolveSheetConfiguration(config) {
    let sheetWIn = config.sheetWidthIn;
    let sheetHIn = config.sheetHeightIn;

    if (config.sheetOrientation === "portrait") {
      sheetWIn = Math.min(config.sheetWidthIn, config.sheetHeightIn);
      sheetHIn = Math.max(config.sheetWidthIn, config.sheetHeightIn);
    } else if (config.sheetOrientation === "landscape") {
      sheetWIn = Math.max(config.sheetWidthIn, config.sheetHeightIn);
      sheetHIn = Math.min(config.sheetWidthIn, config.sheetHeightIn);
    }

    let capacity = maxPlacementsForSheet(
      sheetWIn,
      sheetHIn,
      config.cutWidthIn,
      config.cutHeightIn,
      config.gapHorizontalIn,
      config.gapVerticalIn
    );
    const swappedCapacity = maxPlacementsForSheet(
      sheetHIn,
      sheetWIn,
      config.cutWidthIn,
      config.cutHeightIn,
      config.gapHorizontalIn,
      config.gapVerticalIn
    );

    let bestFitSwapped = false;
    if (config.bestFit && swappedCapacity > capacity) {
      const priorWidth = sheetWIn;
      sheetWIn = sheetHIn;
      sheetHIn = priorWidth;
      capacity = swappedCapacity;
      bestFitSwapped = true;
    }

    return {
      sheetWIn,
      sheetHIn,
      capacity,
      bestFitSwapped,
      actualOrientation: orientationFromDimensions(sheetWIn, sheetHIn)
    };
  }

  function planLayout(sheetWIn, sheetHIn, cutWIn, cutHIn, gapHIn, gapVIn, requiredCount) {
    const outerMarginIn = 0.125;
    const required = Math.max(1, Number(requiredCount) || 1);
    const availWIn = sheetWIn - outerMarginIn * 2;
    const availHIn = sheetHIn - outerMarginIn * 2;
    if (availWIn <= 0 || availHIn <= 0) return null;

    const { colsMax, rowsMax } = gridFit(availWIn, availHIn, cutWIn, cutHIn, gapHIn, gapVIn);
    const maxPlacements = colsMax * rowsMax;
    if (maxPlacements < required) return null;

    let cols = Math.min(colsMax, required);
    let rowsNeeded = Math.ceil(required / cols);
    while (rowsNeeded > rowsMax && cols > 1) {
      cols -= 1;
      rowsNeeded = Math.ceil(required / cols);
    }
    if (rowsNeeded > rowsMax) return null;

    return {
      cols,
      rows: rowsNeeded
    };
  }

  function maxPlacementsForSheet(sheetWIn, sheetHIn, cutWIn, cutHIn, gapHIn, gapVIn) {
    const outerMarginIn = 0.125;
    const availWIn = sheetWIn - outerMarginIn * 2;
    const availHIn = sheetHIn - outerMarginIn * 2;
    if (availWIn <= 0 || availHIn <= 0 || cutWIn <= 0 || cutHIn <= 0) return 0;

    const { colsMax, rowsMax } = gridFit(availWIn, availHIn, cutWIn, cutHIn, gapHIn, gapVIn);
    return colsMax * rowsMax;
  }

  function gridFit(availWIn, availHIn, cellWIn, cellHIn, gapHIn, gapVIn) {
    const colsMax = Math.max(1, Math.floor((availWIn + gapHIn) / (cellWIn + gapHIn)));
    const rowsMax = Math.max(1, Math.floor((availHIn + gapVIn) / (cellHIn + gapVIn)));
    return { colsMax, rowsMax };
  }

  function buildAdjustmentSummary(config, resolvedSheet) {
    const parts = [];
    if (resolvedSheet.bestFitSwapped) {
      parts.push(`best fit swapped the sheet to ${resolvedSheet.actualOrientation}`);
    } else if (config.bestFit) {
      parts.push("best fit kept the current sheet orientation");
    }
    if (config.duplex) {
      parts.push(`duplex handling on ${config.bindingEdge.toLowerCase()} binding`);
    }
    if (config.artRotation !== "None") {
      parts.push(`${String(config.artRotation).toLowerCase()} 180 degree art rotation`);
    }
    if (config.imageShiftXIn || config.imageShiftYIn) {
      parts.push(`image shift ${formatSigned(config.imageShiftXIn)} x, ${formatSigned(config.imageShiftYIn)} y`);
    }
    if (config.autoCorrectArtOrientation) {
      parts.push("auto-correct orientation enabled");
    }
    return parts.length ? parts.join(" • ") : "No extra adjustments enabled.";
  }

  function buildNoteSummary(config, layout, resolvedSheet, sourceMode, sourceLabel, lastResult) {
    if (!layout) {
      return "The current sizes do not fit within the sheet after the fixed 0.125 inch outer margins are applied.";
    }

    const capacity = layout.cols * layout.rows;
    const modeSummary =
      config.mode === "cutAndStack"
        ? "Cut-and-stack fills position 1 across all sheets before moving to position 2."
        : "Repeat mode fills every position on a sheet with the current source page.";
    const sourceSummary =
      sourceMode === "inbox"
        ? "Drop PDFs into LocalImposition/inbox, refresh the list, and pick the one you want to impose."
        : `Current upload source: ${sourceLabel}.`;
    const outputSummary = lastResult
      ? `The last run produced ${lastResult.outputSheetCount} imposed sheet(s) from ${lastResult.pageCount} source page(s).`
      : "The backend will determine the source page count when the job starts.";
    return `${modeSummary} This layout yields ${capacity} positions per sheet. ${sourceSummary} ${outputSummary} Final sheet orientation is ${resolvedSheet.actualOrientation}.`;
  }

  function orientationFromDimensions(width, height) {
    if (width <= 0 || height <= 0) return "square";
    if (Math.abs(width - height) <= 0.0001) return "square";
    return height > width ? "portrait" : "landscape";
  }

  function readPositiveNumber(input, label) {
    const value = Number.parseFloat(input.value);
    if (!Number.isFinite(value) || value <= 0) {
      throw new Error(`${label} must be a number greater than zero.`);
    }
    return value;
  }

  function readNonNegativeNumber(input, label) {
    const value = Number.parseFloat(input.value);
    if (!Number.isFinite(value) || value < 0) {
      throw new Error(`${label} must be zero or greater.`);
    }
    return value;
  }

  function readSignedNumber(input) {
    const value = Number.parseFloat(input.value);
    return Number.isFinite(value) ? value : 0;
  }

  function setBusy(isBusy) {
    state.busy = isBusy;
    elements.generateButton.disabled = isBusy;
    elements.refreshInbox.disabled = isBusy;
    elements.generateButton.textContent = isBusy ? "Processing..." : "Generate imposed PDF";
  }

  function setStatus(message, stateName) {
    elements.status.textContent = message;
    elements.status.dataset.state = stateName;
  }

  function formatFileSize(size) {
    const units = ["B", "KB", "MB", "GB", "TB"];
    let value = size;
    let unitIndex = 0;
    while (value >= 1024 && unitIndex < units.length - 1) {
      value /= 1024;
      unitIndex += 1;
    }
    return `${value.toFixed(value >= 10 || unitIndex === 0 ? 0 : 1)} ${units[unitIndex]}`;
  }

  function formatNumber(value) {
    return Number(value.toFixed(3)).toString();
  }

  function formatSigned(value) {
    const num = Number(value) || 0;
    const prefix = num >= 0 ? "+" : "";
    return `${prefix}${formatNumber(num)}`;
  }
})();

function initUploadForm(options) {
  const {
    formId = "uploadForm",
    dropZoneId = "dropZone",
    fileInputId = "fileInput",
    fileInfoId = "fileInfo",
    fileNameId = "fileName",
    removeFileId = "removeFile",
    submitBtnId = "submitBtn",
    errorBannerId = "errorBanner",
    stepsSelector = ".step",
  } = options;

  const uploadForm = document.getElementById(formId);
  const dropZone = document.getElementById(dropZoneId);
  const fileInput = document.getElementById(fileInputId);
  const fileInfo = document.getElementById(fileInfoId);
  const fileName = document.getElementById(fileNameId);
  const removeFile = document.getElementById(removeFileId);
  const submitBtn = document.getElementById(submitBtnId);
  const errorBanner = document.getElementById(errorBannerId);
  const steps = document.querySelectorAll(stepsSelector);

  function setStep(index) {
    steps.forEach((step, i) => {
      step.classList.remove("active", "done");
      if (i < index) step.classList.add("done");
      if (i === index) step.classList.add("active");
    });
  }

  function showError(message) {
    if (!errorBanner) return;
    errorBanner.textContent = message;
    errorBanner.classList.add("visible");
  }

  function hideError() {
    if (!errorBanner) return;
    errorBanner.classList.remove("visible");
    errorBanner.textContent = "";
  }

  function showFile(file) {
    hideError();
    fileName.textContent = file.name;
    fileInfo.classList.add("visible");
    submitBtn.disabled = false;
    setStep(1);
  }

  function clearFile() {
    fileInput.value = "";
    fileInfo.classList.remove("visible");
    submitBtn.disabled = true;
    setStep(0);
  }

  fileInput.addEventListener("change", () => {
    if (fileInput.files[0]) showFile(fileInput.files[0]);
  });

  removeFile.addEventListener("click", clearFile);

  dropZone.addEventListener("dragover", (e) => {
    e.preventDefault();
    dropZone.classList.add("dragover");
  });
  dropZone.addEventListener("dragleave", () => dropZone.classList.remove("dragover"));
  dropZone.addEventListener("drop", (e) => {
    e.preventDefault();
    dropZone.classList.remove("dragover");
    const file = e.dataTransfer.files[0];
    if (file && (file.name.endsWith(".xls") || file.name.endsWith(".xlsx"))) {
      const dt = new DataTransfer();
      dt.items.add(file);
      fileInput.files = dt.files;
      showFile(file);
    }
  });

  uploadForm.addEventListener("submit", async (e) => {
    e.preventDefault();
    hideError();
    submitBtn.classList.add("loading");
    submitBtn.disabled = true;
    setStep(2);

    try {
      const formData = new FormData(uploadForm);
      const resp = await fetch(uploadForm.action || window.location.pathname, {
        method: "POST",
        body: formData,
      });

      if (!resp.ok) {
        const text = await resp.text();
        throw new Error(text || `Request failed (${resp.status})`);
      }

      const blob = await resp.blob();
      let filename = "output.xml";
      const disposition = resp.headers.get("Content-Disposition");
      if (disposition) {
        const match = disposition.match(/filename\*?=(?:UTF-8'')?"?([^";]+)"?/i);
        if (match) filename = decodeURIComponent(match[1]);
      }

      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(url);

      // brief "done" flash before resetting so the user sees success
      setStep(3 <= steps.length - 1 ? 2 : 2);
      steps.forEach((step) => step.classList.add("done"));
      setTimeout(() => {
        submitBtn.classList.remove("loading");
        clearFile();
      }, 900);
    } catch (err) {
      submitBtn.classList.remove("loading");
      submitBtn.disabled = false;
      setStep(1);
      showError(err.message || "Something went wrong. Please try again.");
    }
  });
}

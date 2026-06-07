const config = window.FATEH_TIME_CONFIG || {};
const form = document.querySelector("#timeForm");
const message = document.querySelector("#message");
const clockText = document.querySelector("#clockText");
const todayText = document.querySelector("#todayText");
const locationStatus = document.querySelector("#locationStatus");
const refreshLocation = document.querySelector("#refreshLocation");
const actionButtons = Array.from(document.querySelectorAll("[data-action]"));
const workerNameInput = document.querySelector("#workerName");
const workerPinInput = document.querySelector("#workerPin");
const actions = document.querySelector(".actions");

let currentPosition = null;
let pendingAction = "CHECK_IN";
let allowedAction = null;
let busy = false;
let statusTimer = null;
let locationTimer = null;

updateActionButtons();

function updateClock() {
  const now = new Date();
  todayText.textContent = now.toLocaleDateString(undefined, {
    weekday: "long",
    month: "short",
    day: "numeric",
    year: "numeric",
  });
  clockText.textContent = now.toLocaleTimeString(undefined, {
    hour: "numeric",
    minute: "2-digit",
  });
}

function setMessage(text, type = "") {
  message.textContent = text;
  message.className = `message ${type}`.trim();
}

function updateActionButtons() {
  const hasPin = workerPinInput.value.trim() !== "";
  const hasWorker = workerNameInput.value.trim() !== "";
  let visibleCount = 0;
  actionButtons.forEach((button) => {
    const isAllowed = !allowedAction || button.dataset.action === allowedAction;
    const shouldHide = !hasPin || !hasWorker || (allowedAction && !isAllowed);
    button.classList.toggle("is-hidden", shouldHide);
    button.disabled = busy || shouldHide || !isAllowed;
    if (!button.classList.contains("is-hidden")) {
      visibleCount += 1;
    }
  });
  actions.classList.toggle("single", visibleCount === 1);
}

function setBusy(nextBusy) {
  busy = nextBusy;
  updateActionButtons();
  refreshLocation.disabled = nextBusy;
}

function setNextAction(nextAction) {
  if (!nextAction) {
    return;
  }

  allowedAction = nextAction;
  updateActionButtons();
  pendingAction = nextAction;
}

function queueStatusCheck() {
  clearTimeout(statusTimer);
  allowedAction = null;
  workerNameInput.value = "";
  updateActionButtons();
  statusTimer = setTimeout(checkWorkerStatus, 450);
}

async function checkWorkerStatus() {
  const scriptUrl = (config.scriptUrl || "").trim();
  const pin = workerPinInput.value.trim();

  if (!scriptUrl || !pin) {
    allowedAction = null;
    workerNameInput.value = "";
    updateActionButtons();
    setMessage("Enter worker PIN to load current status.");
    return;
  }

  try {
    const response = await jsonpRequest(scriptUrl, {
      action: "CHECK_STATUS",
      pin,
    });

    if (!response || response.ok !== true) {
      allowedAction = null;
      workerNameInput.value = "";
      updateActionButtons();
      if (
        response?.message === "Choose check in or check out." ||
        response?.message === "Name and PIN are required."
      ) {
        setMessage("Apps Script deployment is old. Deploy a new Web App version to use PIN lookup.", "bad");
      } else {
        setMessage(response?.message || "Enter a valid name and PIN.", "bad");
      }
      return;
    }

    workerNameInput.value = response.name || "";
    setNextAction(response.nextAction);
    setMessage(`${response.name}: ${response.message}`, "");
  } catch {
    allowedAction = null;
  }
}

function requestLocation() {
  clearTimeout(locationTimer);
  if (!navigator.geolocation) {
    locationStatus.textContent = "Location is not supported on this device";
    setMessage("GPS is unavailable on this device. You can continue without location.");
    return;
  }

  locationStatus.textContent = "Getting current location...";
  locationTimer = setTimeout(() => {
    if (!currentPosition) {
      locationStatus.textContent = "Location unavailable, continue without GPS";
      setMessage("GPS did not respond. You can still submit without location.");
    }
  }, 8000);

  navigator.geolocation.getCurrentPosition(
    (position) => {
      clearTimeout(locationTimer);
      currentPosition = position;
      const accuracy = Math.round(position.coords.accuracy);
      locationStatus.textContent = `Ready, accurate within about ${accuracy} m`;
    },
    (error) => {
      clearTimeout(locationTimer);
      currentPosition = null;
      const reason =
        error.code === error.PERMISSION_DENIED
          ? "Location permission was denied"
          : "Could not get current location";
      locationStatus.textContent = reason;
      setMessage("GPS is not available right now. You can continue without location.");
    },
    {
      enableHighAccuracy: true,
      timeout: 8000,
      maximumAge: 300000,
    },
  );
}

function jsonpRequest(url, params) {
  return new Promise((resolve, reject) => {
    const callbackName = `fatehTime_${Date.now()}_${Math.random().toString(16).slice(2)}`;
    const script = document.createElement("script");
    const searchParams = new URLSearchParams({
      ...params,
      callback: callbackName,
    });

    window[callbackName] = (response) => {
      cleanup();
      resolve(response);
    };

    function cleanup() {
      script.remove();
      delete window[callbackName];
    }

    script.onerror = () => {
      cleanup();
      reject(new Error("Could not reach Google Sheet script."));
    };

    script.src = `${url}?${searchParams.toString()}`;
    document.body.appendChild(script);
  });
}

async function submitTime(action) {
  if (allowedAction && action !== allowedAction) {
    setMessage(
      allowedAction === "CHECK_IN" ? "This worker needs to check in first." : "This worker needs to check out first.",
      "bad",
    );
    return;
  }

  const scriptUrl = (config.scriptUrl || "").trim();
  if (!scriptUrl) {
    setMessage("Add your Google Apps Script Web App URL in config.js first.", "bad");
    return;
  }

  if (!form.reportValidity()) {
    return;
  }

  const now = new Date();
  const formData = new FormData(form);
  const coords = currentPosition?.coords || null;

  setBusy(true);
  setMessage(action === "CHECK_IN" ? "Saving check in..." : "Saving check out...");

  try {
    const response = await jsonpRequest(scriptUrl, {
      action,
      pin: String(formData.get("pin") || "").trim(),
      name: String(formData.get("name") || "").trim(),
      job: String(formData.get("job") || "").trim(),
      notes: String(formData.get("notes") || "").trim(),
      clientLocalTime: now.toLocaleString(),
      clientIsoTime: now.toISOString(),
      timezone: Intl.DateTimeFormat().resolvedOptions().timeZone || "",
      latitude: coords ? coords.latitude.toFixed(7) : "",
      longitude: coords ? coords.longitude.toFixed(7) : "",
      accuracyMeters: coords ? Math.round(coords.accuracy) : "",
      userAgent: navigator.userAgent,
    });

    if (!response || response.ok !== true) {
      if (response?.nextAction) {
        setNextAction(response.nextAction);
      }
      throw new Error(response?.message || "The entry was not saved.");
    }

    const actionText = action === "CHECK_IN" ? "checked in" : "checked out";
    setMessage(`${response.name} ${actionText} at ${response.time}.`, "good");
    setNextAction(response.nextAction);
    form.querySelector("#jobName").value = "";
    form.querySelector("#notes").value = "";
    requestLocation();
  } catch (error) {
    setMessage(error.message || "Something went wrong. Try again.", "bad");
  } finally {
    setBusy(false);
  }
}

actionButtons.forEach((button) => {
  button.addEventListener("click", () => {
    if (allowedAction && button.dataset.action !== allowedAction) {
      return;
    }
    pendingAction = button.dataset.action;
  });
});

form.addEventListener("submit", (event) => {
  event.preventDefault();
  submitTime(pendingAction);
});

refreshLocation.addEventListener("click", requestLocation);
workerPinInput.addEventListener("input", queueStatusCheck);

updateClock();
setInterval(updateClock, 1000);
requestLocation();

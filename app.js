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

let currentPosition = null;
let pendingAction = "CHECK_IN";
let allowedAction = null;
let busy = false;
let statusTimer = null;

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
  actionButtons.forEach((button) => {
    button.disabled = busy || (allowedAction && button.dataset.action !== allowedAction);
  });
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
  statusTimer = setTimeout(checkWorkerStatus, 450);
}

async function checkWorkerStatus() {
  const scriptUrl = (config.scriptUrl || "").trim();
  const name = workerNameInput.value.trim();
  const pin = workerPinInput.value.trim();

  if (!scriptUrl || !name || !pin) {
    allowedAction = null;
    updateActionButtons();
    return;
  }

  try {
    const response = await jsonpRequest(scriptUrl, {
      action: "CHECK_STATUS",
      name,
      pin,
    });

    if (!response || response.ok !== true) {
      allowedAction = null;
      updateActionButtons();
      return;
    }

    setNextAction(response.nextAction);
    setMessage(`${response.name}: ${response.message}`, "");
  } catch {
    allowedAction = null;
  }
}

function requestLocation() {
  if (!navigator.geolocation) {
    locationStatus.textContent = "Location is not supported on this device";
    setMessage("This device/browser cannot read GPS location.", "bad");
    return;
  }

  locationStatus.textContent = "Getting current location...";
  navigator.geolocation.getCurrentPosition(
    (position) => {
      currentPosition = position;
      const accuracy = Math.round(position.coords.accuracy);
      locationStatus.textContent = `Ready, accurate within about ${accuracy} m`;
    },
    (error) => {
      currentPosition = null;
      const reason =
        error.code === error.PERMISSION_DENIED
          ? "Location permission was denied"
          : "Could not get current location";
      locationStatus.textContent = reason;
      setMessage("Please allow location before checking in or out.", "bad");
    },
    {
      enableHighAccuracy: true,
      timeout: 12000,
      maximumAge: 60000,
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
  const scriptUrl = (config.scriptUrl || "").trim();
  if (!scriptUrl) {
    setMessage("Add your Google Apps Script Web App URL in config.js first.", "bad");
    return;
  }

  if (!currentPosition) {
    setMessage("GPS location is required. Tap refresh and allow location.", "bad");
    requestLocation();
    return;
  }

  if (!form.reportValidity()) {
    return;
  }

  const now = new Date();
  const formData = new FormData(form);
  const coords = currentPosition.coords;

  setBusy(true);
  setMessage(action === "CHECK_IN" ? "Saving check in..." : "Saving check out...");

  try {
    const response = await jsonpRequest(scriptUrl, {
      action,
      name: String(formData.get("name") || "").trim(),
      pin: String(formData.get("pin") || "").trim(),
      job: String(formData.get("job") || "").trim(),
      notes: String(formData.get("notes") || "").trim(),
      clientLocalTime: now.toLocaleString(),
      clientIsoTime: now.toISOString(),
      timezone: Intl.DateTimeFormat().resolvedOptions().timeZone || "",
      latitude: coords.latitude.toFixed(7),
      longitude: coords.longitude.toFixed(7),
      accuracyMeters: Math.round(coords.accuracy),
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
workerNameInput.addEventListener("input", queueStatusCheck);
workerPinInput.addEventListener("input", queueStatusCheck);

updateClock();
setInterval(updateClock, 1000);
requestLocation();

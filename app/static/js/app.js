const workTypeSelect = document.getElementById("workType");
const stationSelect = document.getElementById("station");
const equipmentSelect = document.getElementById("equipment");
const generateButton = document.getElementById("generateButton");
const statusBox = document.getElementById("status");

function setStatus(message, type = "") {
  statusBox.textContent = message;
  statusBox.className = `status ${type}`.trim();
}


function fillSelect(select, options, placeholder) {
  select.innerHTML = "";
  const emptyOption = document.createElement("option");
  emptyOption.value = "";
  emptyOption.textContent = placeholder;
  select.appendChild(emptyOption);
  for (const optionValue of options) {
    const option = document.createElement("option");
    option.value = optionValue;
    option.textContent = optionValue;
    select.appendChild(option);
  }
}


function updateGenerateState() {
  generateButton.disabled = !(workTypeSelect.value && stationSelect.value && equipmentSelect.value);
}


async function loadStations() {
  try {
    const response = await fetch("/api/stations");
    if (!response.ok) throw new Error((await response.json()).detail || "Не удалось загрузить список КС.");
    const data = await response.json();
    fillSelect(stationSelect, data.stations, "Выберите КС");
    stationSelect.disabled = false;
  } catch (error) {
    fillSelect(stationSelect, [], "Ошибка загрузки КС");
    setStatus(error.message, "error");
  }
}


async function loadEquipment(station) {
  equipmentSelect.disabled = true;
  fillSelect(equipmentSelect, [], "Загрузка списка ГПА…");
  try {
    const response = await fetch(`/api/equipment?station=${encodeURIComponent(station)}`);
    if (!response.ok) throw new Error((await response.json()).detail || "Не удалось загрузить список ГПА.");
    const data = await response.json();
    fillSelect(equipmentSelect, data.equipment, "Выберите ГПА");
    equipmentSelect.disabled = false;
  } catch (error) {
    fillSelect(equipmentSelect, [], "Ошибка загрузки ГПА");
    setStatus(error.message, "error");
  }
}


stationSelect.addEventListener("change", async () => {
  equipmentSelect.value = "";
  updateGenerateState();
  if (stationSelect.value) await loadEquipment(stationSelect.value);
  else {
    equipmentSelect.disabled = true;
    fillSelect(equipmentSelect, [], "Сначала выберите КС");
  }
});

workTypeSelect.addEventListener("change", updateGenerateState);
equipmentSelect.addEventListener("change", updateGenerateState);


generateButton.addEventListener("click", async () => {
  if (!workTypeSelect.value || !stationSelect.value || !equipmentSelect.value) return;
  generateButton.disabled = true;
  generateButton.textContent = "Формирование…";
  setStatus("Формируем файл, пожалуйста, подождите.");
  try {
    const response = await fetch("/api/generate", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        work_type: workTypeSelect.value,
        station: stationSelect.value,
        equipment: equipmentSelect.value,
      }),
    });
    if (!response.ok) {
      let detail = "Не удалось сформировать ведомость.";
      try { detail = (await response.json()).detail || detail; } catch (_) {}
      throw new Error(detail);
    }
    const blob = await response.blob();
    const disposition = response.headers.get("Content-Disposition") || "";
    const filenameMatch = disposition.match(/filename\*=UTF-8''([^;]+)|filename="?([^";]+)"?/i);
    const filename = filenameMatch ? decodeURIComponent(filenameMatch[1] || filenameMatch[2]) : "Ведомость_ГПА.xlsx";
    const downloadUrl = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = downloadUrl;
    anchor.download = filename;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    URL.revokeObjectURL(downloadUrl);
    setStatus("Файл сформирован и скачан.", "success");
  } catch (error) {
    setStatus(error.message, "error");
  } finally {
    generateButton.textContent = "Сформировать ведомость";
    updateGenerateState();
  }
});

loadStations();

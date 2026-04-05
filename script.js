class Confirm {
  constructor({ id, room, city }) {
    this.id = id;
    this.room = room;
    this.city = city;
  }

  log() {
    console.log(`Confirm: ${this.id}, ${this.room}, ${this.city}`);
  }
}

function resetSelectedFile() {
  const fileInput = document.getElementById("file");
  const fileLabel = document.getElementById("fileLabel");

  if (fileInput) {
    fileInput.value = "";
  }

  if (fileLabel) {
    fileLabel.textContent = "📄 Upload File Excel (.xlsx / .xls)";
  }
}

let activeAlertResolver = null;
let lastFocusedElement = null;

function hideAlertModal() {
  const modal = document.getElementById("alertModal");

  if (!modal) return;

  modal.classList.remove("is-visible");
  modal.setAttribute("aria-hidden", "true");

  if (activeAlertResolver) {
    const resolve = activeAlertResolver;
    activeAlertResolver = null;
    resolve();
  }

  if (lastFocusedElement && typeof lastFocusedElement.focus === "function") {
    lastFocusedElement.focus();
    lastFocusedElement = null;
  }
}

function showAlertModal(message, title = "Alert", variant = "error") {
  const modal = document.getElementById("alertModal");
  const cardElement = document.getElementById("alertCard");
  const iconElement = document.getElementById("alertModalIcon");
  const titleElement = document.getElementById("alertModalTitle");
  const messageElement = document.getElementById("alertModalMessage");
  const buttonElement = document.getElementById("alertModalButton");

  if (
    !modal ||
    !cardElement ||
    !iconElement ||
    !titleElement ||
    !messageElement ||
    !buttonElement
  ) {
    return Promise.resolve();
  }

  if (activeAlertResolver) {
    activeAlertResolver();
    activeAlertResolver = null;
  }

  lastFocusedElement = document.activeElement;
  cardElement.dataset.state = variant;
  iconElement.textContent = variant === "success" ? "✓" : "!";
  titleElement.textContent = title;
  messageElement.textContent = message;
  modal.classList.add("is-visible");
  modal.setAttribute("aria-hidden", "false");

  buttonElement.focus();

  return new Promise((resolve) => {
    activeAlertResolver = resolve;
  });
}

document.addEventListener("click", (event) => {
  const modal = document.getElementById("alertModal");

  if (modal && event.target === modal) {
    hideAlertModal();
  }
});

document.addEventListener("keydown", (event) => {
  if (event.key === "Escape") {
    hideAlertModal();
  }
});

async function exportExcel() {
  try {
    const confirmList = await convertToComfirmList();

    if (!confirmList || confirmList.length === 0) {
      await showAlertModal(
        "The selected file contains invalid data.",
        "Invalid Data",
      );
      resetSelectedFile();
      return;
    }

    const rows = buildExcelRows(confirmList);
    const worksheet = XLSX.utils.aoa_to_sheet(rows);
    formatDefaultWorksheet(worksheet);

    worksheet["!merges"] = [
      {
        s: { r: 0, c: 0 },
        e: { r: 0, c: 5 },
      },
    ];

    const titleCell = worksheet["A1"];

    titleCell.s = {
      font: {
        bold: true,
        sz: 20,
      },
      alignment: {
        horizontal: "center",
        vertical: "center",
      },
    };

    let rowStart = null;
    rows.forEach((row, index) => {
      if (index <= 1) return;

      if (row[0] === "") {
        // Format room number
        const addr = "B" + (index + 1);
        worksheet[addr].s = {
          font: {
            sz: 20,
            bold: true,
          },
          alignment: {
            horizontal: "center",
            vertical: "center",
          },
        };

        if (rowStart === null) rowStart = index;
        return;
      }

      const rowEnd = index - 1;

      if (rowEnd > rowStart && rowStart !== null) {
        worksheet["!merges"].push(
          {
            s: { r: rowStart, c: 0 },
            e: { r: rowEnd, c: 0 },
          },
          {
            s: { r: rowStart, c: 3 },
            e: { r: rowEnd, c: 3 },
          },
        );
      }

      worksheet["!merges"].push({
        s: { r: index, c: 2 },
        e: { r: index, c: 5 },
      });

      rowStart = null;
    });

    worksheet["!rows"] = rows.map((row, index) => {
      const roomCell = row[1];

      if (index === 0) {
        return { hpt: 60 };
      }

      if (roomCell && !isNaN(roomCell) && row[0] === "") {
        return { hpt: 45 };
      }

      return { hpt: 20 };
    });

    worksheet["!cols"] = [
      { wch: 20 },
      { wch: 12 },
      { wch: 7 },
      { wch: 12 },
      { wch: 20 },
      { wch: 12 },
    ];

    addBorderAllCells(worksheet);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");

    const today = getTodayDDMMYYYY();
    const filename = `CHECKOUT_${today.replace(/\//g, "-")}.xlsx`;

    XLSX.writeFile(workbook, filename);
    await showAlertModal(
      `Saved ${filename} successfully.`,
      "Export complete",
      "success",
    );
    resetSelectedFile();
  } catch (error) {
    if (error?.silent) {
      await showAlertModal(
        error.message || "Something went wrong.",
        error.title || "Alert",
      );
      return;
    }

    console.error(error);
    await showAlertModal(
      "The Excel file could not be processed. Please check the file and try again.",
      "Export failed",
    );
  }
}

function formatDefaultWorksheet(worksheet) {
  Object.keys(worksheet).forEach((cellAddress) => {
    if (cellAddress.startsWith("!")) return;

    const cell = worksheet[cellAddress];
    if (!cell || cell.v === undefined) return;

    cell.s = cell.s || {};
    cell.s = {
      ...cell.s,
      font: {
        sz: 12,
      },
      alignment: {
        horizontal: "center",
        vertical: "center",
      },
    };
  });
}

function addBorderAllCells(worksheet) {
  const borderStyle = {
    top: { style: "thin" },
    bottom: { style: "thin" },
    left: { style: "thin" },
    right: { style: "thin" },
  };

  Object.keys(worksheet).forEach((cellAddress) => {
    if (cellAddress.startsWith("!")) return;

    const cell = worksheet[cellAddress];
    if (!cell || cell.v === undefined) return;

    cell.s = cell.s || {};
    cell.s.border = borderStyle;
  });
}

function convertToComfirmList() {
  return new Promise((resolve, reject) => {
    const file = document.getElementById("file").files[0];

    if (!file) {
      reject({
        silent: true,
        title: "No file selected",
        message: "Please choose an Excel file before exporting.",
      });
      return;
    }

    const reader = new FileReader();

    reader.onload = (event) => {
      try {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, {
          header: 1,
          defval: "",
        });

        const confirmList = sortConfirms(parseConfirms(rows));

        console.log("Total confirms:", confirmList.length);
        confirmList.forEach((confirm) => confirm.log());

        resolve(confirmList);
      } catch (error) {
        reject(error);
      }
    };

    reader.onerror = () => {
      reject(new Error("Failed to read the selected file."));
    };

    reader.readAsArrayBuffer(file);
  });
}

function parseConfirms(rows) {
  const confirms = [];
  let current = null;

  rows.forEach((row) => {
    const firstCell = row[0]?.toString().trim();

    if (firstCell.startsWith("Confirm")) {
      if (current) {
        current.room.sort((a, b) => Number(a) - Number(b));
        confirms.push(current);
      }

      current = {
        id: row[1],
        city: row[2],
        room: [],
      };
      return;
    }

    if (current && firstCell && !isNaN(firstCell)) {
      current.room.push(firstCell);
    }
  });

  if (current) confirms.push(current);

  return confirms.map((confirm) => new Confirm(confirm));
}

function sortConfirms(confirmList) {
  return confirmList.sort((a, b) => {
    const roomA = Number(a.room[0] ?? Infinity);
    const roomB = Number(b.room[0] ?? Infinity);
    return roomA - roomB;
  });
}

function buildExcelRows(confirmList) {
  const rows = [];
  const today = getTodayDDMMYYYY();

  rows.push(["CHECK OUT LIST NGÀY " + today, "", "", "", "", ""]);
  rows.push(["BK.No", "Room", "Key", "R/C", "Minibar", "Other"]);

  confirmList.forEach((confirm) => {
    rows.push(["Confirm Num:", confirm.id, confirm.city, "", "", ""]);

    confirm.room.forEach((room) => {
      rows.push(["", room, "", "", "", ""]);
    });
  });

  return rows;
}

function getTodayDDMMYYYY() {
  const now = new Date();
  const d = String(now.getDate()).padStart(2, "0");
  const m = String(now.getMonth() + 1).padStart(2, "0");
  const y = now.getFullYear();
  return `${d}/${m}/${y}`;
}

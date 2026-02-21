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

async function exportExcel() {
  const confirmList = await convertToComfirmList();

  if (!confirmList || confirmList.length === 0) {
    alert("File chưa có dữ liệu để xuất");
    return;
  }

  const rows = buildExcelRows(confirmList);
  const worksheet = XLSX.utils.aoa_to_sheet(rows);

  worksheet["!merges"] = [
    {
      s: { r: 0, c: 0 },
      e: { r: 0, c: 5 },
    },
  ];

  worksheet["!rows"] = rows.map((row, index) => {
    const roomCell = row[1];

    if (roomCell && !isNaN(roomCell) && row[0] == "") {
      return { hpt: 50 };
    }

    return { hpt: 20 };
  });

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");

  const today = getTodayDDMMYYYY();
  const filename = `CHECKOUT_${today.replace(/\//g, "-")}.xlsx`;

  XLSX.writeFile(workbook, filename);
}

function convertToComfirmList() {
  return new Promise((resolve, reject) => {
    const file = document.getElementById("file").files[0];
    if (!file) {
      alert("Hãy chon file trước khi xuất");
      return reject("No file");
    }

    const reader = new FileReader();

    reader.onload = (e) => {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array" });

      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        defval: "",
      });

      const confirmList = sortConfirms(parseConfirms(rows));

      console.log("Total confirms:", confirmList.length);
      confirmList.forEach((c) => c.log());

      resolve(confirmList);
    };

    reader.onerror = reject;
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

  return confirms.map((c) => new Confirm(c));
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

const express = require("express");
const multer = require("multer");
const { CloudinaryStorage } = require("multer-storage-cloudinary");
const cloudinary = require("cloudinary").v2;
const ExcelJS = require("exceljs");
const { Document, Packer, Paragraph, TextRun } = require("docx");

const app = express();
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// -------------------- Cloudinary Config --------------------
cloudinary.config({
  cloud_name: process.env.CLOUDINARY_CLOUD_NAME,
  api_key: process.env.CLOUDINARY_API_KEY,
  api_secret: process.env.CLOUDINARY_API_SECRET,
});

const storage = new CloudinaryStorage({
  cloudinary: cloudinary,
  params: {
    folder: "laporan_foto",
    allowed_formats: ["jpg", "png", "jpeg"],
  },
});

const upload = multer({ storage });

// -------------------- Data memori --------------------
let laporan = [];

// -------------------- Helper --------------------
function formatTanggal(tanggal) {
  if (!tanggal) return "-";
  return new Date(tanggal).toLocaleString("id-ID", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
  });
}

// -------------------- Halaman Upload --------------------
app.get("/", (req, res) => {
  res.send(`
  <html>
  <head>
    <title>Upload Laporan</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
  </head>
  <body class="bg-light p-4">
    <div class="container">
      <h2 class="mb-4">➕ Form Upload Laporan</h2>
      <form action="/upload" method="post" enctype="multipart/form-data" class="card p-4 shadow-sm">
        <div class="mb-3">
          <label class="form-label">Foto</label>
          <input type="file" name="foto" class="form-control" required>
        </div>
        <div class="mb-3">
          <label class="form-label">Gardu</label>
          <input type="text" name="gardu" class="form-control" required>
        </div>
        <div class="mb-3">
          <label class="form-label">Kerusakan</label>
          <input type="text" name="kerusakan" class="form-control" required>
        </div>
        <div class="mb-3">
          <label class="form-label">Waktu Kerusakan</label>
          <input type="datetime-local" name="waktuKerusakan" class="form-control" required>
        </div>
        <div class="mb-3">
          <label class="form-label">Perbaikan</label>
          <input type="text" name="perbaikan" class="form-control" required>
        </div>
        <div class="mb-3">
          <label class="form-label">Waktu Selesai</label>
          <input type="datetime-local" name="selesai" class="form-control" required>
        </div>
        <div class="mb-3">
          <label class="form-label">Status</label>
          <select name="status" class="form-select">
            <option value="Selesai">Selesai</option>
            <option value="Proses">Proses</option>
            <option value="Pending">Pending</option>
          </select>
        </div>
        <div class="mb-3">
          <label class="form-label">By</label>
          <input type="text" name="by" class="form-control" required>
        </div>
        <button type="submit" class="btn btn-primary">Upload</button>
      </form>
    </div>
  </body>
  </html>
  `);
});

// -------------------- Upload Endpoint --------------------
app.post("/upload", upload.single("foto"), (req, res) => {
  const { gardu, kerusakan, perbaikan, selesai, status, by, waktuKerusakan } = req.body;
  const data = {
    id: Date.now(),
    fotoURL: req.file.path,
    filename: req.file.filename,
    originalname: req.file.originalname,
    gardu,
    kerusakan,
    perbaikan,
    selesai: formatTanggal(selesai),
    waktuKerusakan: formatTanggal(waktuKerusakan),
    status,
    by,
    tanggalUpload: new Date().toLocaleString("id-ID"),
  };
  laporan.push(data);
  res.redirect("/laporan");
});

// -------------------- Daftar Laporan --------------------
app.get("/laporan", (req, res) => {
  let html = `
  <html>
  <head>
    <title>Daftar Laporan</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
  </head>
  <body class="bg-light p-4">
    <div class="container">
      <h2 class="mb-4">📋 Daftar Laporan</h2>
      <table class="table table-bordered table-striped table-hover align-middle">
        <thead class="table-primary text-center">
          <tr>
            <th>No</th>
            <th>Foto</th>
            <th>Gardu</th>
            <th>Kerusakan</th>
            <th>Waktu Kerusakan</th>
            <th>Perbaikan</th>
            <th>Waktu Selesai</th>
            <th>Status</th>
            <th>By</th>
            <th>Tanggal Upload</th>
            <th>Edit</th>
            <th>Delete</th>
          </tr>
        </thead>
        <tbody>
  `;

  laporan.forEach((item, index) => {
    html += `
      <tr>
        <td class="text-center">${index + 1}</td>
        <td><img src="${item.fotoURL}" class="img-thumbnail" width="100"/></td>
        <td>${item.gardu}</td>
        <td>${item.kerusakan}</td>
        <td>${item.waktuKerusakan}</td>
        <td>${item.perbaikan}</td>
        <td>${item.selesai}</td>
        <td class="text-center">
          <span class="badge ${item.status === "Selesai" ? "bg-success" : item.status === "Proses" ? "bg-warning text-dark" : "bg-secondary"}">
            ${item.status}
          </span>
        </td>
        <td>${item.by}</td>
        <td>${item.tanggalUpload}</td>
        <td><a href="/edit/${item.id}" class="btn btn-warning btn-sm">Edit</a></td>
        <td><a href="/delete/${item.id}" class="btn btn-danger btn-sm" onclick="return confirm('Yakin hapus?')">Delete</a></td>
      </tr>
    `;
  });

  html += `
        </tbody>
      </table>
      <a href="/" class="btn btn-primary mt-3">➕ Upload Lagi</a>
      <a href="/export/excel" class="btn btn-success mt-3 ms-2">💾 Export Excel</a>
      <a href="/export/word" class="btn btn-secondary mt-3 ms-2">💾 Export Word</a>
    </div>
  </body>
  </html>
  `;
  res.send(html);
});

// -------------------- Delete --------------------
app.get("/delete/:id", (req, res) => {
  const id = parseInt(req.params.id);
  const index = laporan.findIndex(item => item.id === id);
  if (index !== -1) laporan.splice(index, 1);
  res.redirect("/laporan");
});

// -------------------- Edit --------------------
app.get("/edit/:id", (req, res) => {
  const id = parseInt(req.params.id);
  const item = laporan.find(item => item.id === id);
  if (!item) return res.send("Data tidak ditemukan");

  res.send(`
  <html>
  <head>
    <title>Edit Laporan</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
  </head>
  <body class="bg-light p-4">
    <div class="container">
      <h2 class="mb-4">✏️ Edit Laporan</h2>
      <form action="/edit/${id}" method="post" class="card p-4 shadow-sm">
        <div class="mb-3">
          <label class="form-label">Gardu</label>
          <input name="gardu" value="${item.gardu}" class="form-control" required>
        </div>
        <div class="mb-3">
          <label class="form-label">Kerusakan</label>
          <input name="kerusakan" value="${item.kerusakan}" class="form-control" required>
        </div>
        <div class="mb-3">
          <label class="form-label">Waktu Kerusakan</label>
          <input name="waktuKerusakan" type="datetime-local" value="${item.waktuKerusakan}" class="form-control">
        </div>
        <div class="mb-3">
          <label class="form-label">Perbaikan</label>
          <input name="perbaikan" value="${item.perbaikan}" class="form-control">
        </div>
        <div class="mb-3">
          <label class="form-label">Waktu Selesai</label>
          <input name="selesai" type="datetime-local" value="${item.selesai}" class="form-control">
        </div>
        <div class="mb-3">
          <label class="form-label">Status</label>
          <select name="status" class="form-select">
            <option value="Selesai" ${item.status === "Selesai" ? "selected" : ""}>Selesai</option>
            <option value="Proses" ${item.status === "Proses" ? "selected" : ""}>Proses</option>
            <option value="Pending" ${item.status === "Pending" ? "selected" : ""}>Pending</option>
          </select>
        </div>
        <div class="mb-3">
          <label class="form-label">By</label>
          <input name="by" value="${item.by}" class="form-control">
        </div>
        <button type="submit" class="btn btn-success">Update</button>
        <a href="/laporan" class="btn btn-secondary ms-2">Kembali</a>
      </form>
    </div>
  </body>
  </html>
  `);
});

app.post("/edit/:id", (req, res) => {
  const id = parseInt(req.params.id);
  const item = laporan.find(item => item.id === id);
  if (item) {
    const { gardu, kerusakan, perbaikan, selesai, status, by, waktuKerusakan } = req.body;
    Object.assign(item, {
      gardu,
      kerusakan,
      perbaikan,
      selesai: formatTanggal(selesai),
      status,
      by,
      waktuKerusakan: formatTanggal(waktuKerusakan),
    });
  }
  res.redirect("/laporan");
});

// -------------------- Export Excel --------------------
app.get("/export/excel", async (req, res) => {
  const workbook = new ExcelJS.Workbook();
  const ws = workbook.addWorksheet("Laporan");

  ws.columns = [
    { header: "No", key: "no", width: 5 },
    { header: "Foto URL", key: "fotoURL", width: 30 },
    { header: "Gardu", key: "gardu", width: 10 },
    { header: "Kerusakan", key: "kerusakan", width: 30 },
    { header: "Waktu Kerusakan", key: "waktuKerusakan", width: 20 },
    { header: "Perbaikan", key: "perbaikan", width: 30 },
    { header: "Waktu Selesai", key: "selesai", width: 20 },
    { header: "Status", key: "status", width: 15 },
    { header: "By", key: "by", width: 15 },
    { header: "Tanggal Upload", key: "tanggalUpload", width: 25 },
  ];

  laporan.forEach((item, index) => {
    ws.addRow({ no: index + 1, ...item });
  });

  res.setHeader(
    "Content-Type",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  );
  res.setHeader("Content-Disposition", "attachment; filename=laporan.xlsx");
  await workbook.xlsx.write(res);
  res.end();
});

// -------------------- Export Word --------------------
app.get("/export/word", async (req, res) => {
  const doc = new Document({
    sections: [
      {
        properties: {},
        children: laporan.map((item, index) =>
          new Paragraph({
            children: [
              new TextRun(`No: ${index + 1}`),
              new TextRun(`\nFoto URL: ${item.fotoURL}`),
              new TextRun(`\nGardu: ${item.gardu}`),
              new TextRun(`\nKerusakan: ${item.kerusakan}`),
              new TextRun(`\nWaktu Kerusakan: ${item.waktuKerusakan}`),
              new TextRun(`\nPerbaikan: ${item.perbaikan}`),
              new TextRun(`\nWaktu Selesai: ${item.selesai}`),
              new TextRun(`\nStatus: ${item.status}`),
              new TextRun(`\nBy: ${item.by}`),
              new TextRun(`\nTanggal Upload: ${item.tanggalUpload}`),
              new TextRun("\n---------------------------\n"),
            ],
          })
        ),
      },
    ],
  });

  const buffer = await Packer.toBuffer(doc);
  res.setHeader("Content-Disposition", "attachment; filename=laporan.docx");
  res.send(buffer);
});

// -------------------- Jalankan Server --------------------
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server berjalan di http://localhost:${PORT}`));

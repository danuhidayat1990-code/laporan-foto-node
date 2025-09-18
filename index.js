const express = require("express");
const multer = require("multer");
const { CloudinaryStorage } = require("multer-storage-cloudinary");
const cloudinary = require("cloudinary").v2;
const ExcelJS = require("exceljs");
const { Document, Packer, Paragraph, TextRun } = require("docx");
const fs = require("fs");

const app = express();
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// -------------------- Cloudinary Config --------------------
cloudinary.config({
  cloud_name: "ddzctmkri",   // ganti sesuai akun
  api_key: "513619779369171",
  api_secret: "WvqB1vsmRI1IBXFF8NqA9y6EYsM",
});

const storage = new CloudinaryStorage({
  cloudinary: cloudinary,
  params: {
    folder: "laporan_foto",  // nama folder di cloudinary
    allowed_formats: ["jpg", "png", "jpeg"],
  },
});

const upload = multer({ storage });

// -------------------- Data memori --------------------
let laporan = [];

// -------------------- Helper format tanggal --------------------
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
    fotoURL: req.file.path, // URL dari Cloudinary
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

// -------------------- Halaman Daftar --------------------
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
        <td class="text-center"><span class="badge ${item.status === "Selesai" ? "bg-success" : item.status === "Proses" ? "bg-warning text-dark" : "bg-secondary"}">${item.status}</span></td>
        <td>${item.by}</td>
        <td>${item.tanggalUpload}</td>
      </tr>
    `;
  });

  html += `
            </tbody>
          </table>
        </div>
      </body>
    </html>
  `;

  res.send(html);
});

// -------------------- Jalankan server --------------------
app.listen(3000, () => console.log("Server berjalan di http://localhost:3000"));

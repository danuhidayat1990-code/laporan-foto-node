const express = require("express");
const multer = require("multer");
const { CloudinaryStorage } = require("multer-storage-cloudinary");
const cloudinary = require("cloudinary").v2;
const ExcelJS = require("exceljs");
const { Document, Packer, Paragraph, TextRun, Media } = require("docx");
const mongoose = require("mongoose");
const axios = require("axios");
require("dotenv").config();

const app = express();
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// -------------------- MongoDB --------------------
mongoose.connect(process.env.MONGO_URI, {
  useNewUrlParser: true,
  useUnifiedTopology: true,
});

const laporanSchema = new mongoose.Schema({
  fotoURL: String,
  filename: String,
  originalname: String,
  gardu: String,
  kerusakan: String,
  perbaikan: String,
  selesai: String,
  waktuKerusakan: String,
  status: String,
  by: String,
  tanggalUpload: String,
});

const Laporan = mongoose.model("Laporan", laporanSchema);

// -------------------- Cloudinary --------------------
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

// -------------------- Upload Form --------------------
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
app.post("/upload", upload.single("foto"), async (req, res) => {
  const { gardu, kerusakan, perbaikan, selesai, status, by, waktuKerusakan } = req.body;

  const data = new Laporan({
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
  });

  await data.save();
  res.redirect("/laporan");
});

// -------------------- Daftar Laporan --------------------
app.get("/laporan", async (req, res) => {
  const laporan = await Laporan.find().lean();

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
        <td><a href="/edit/${item._id}" class="btn btn-warning btn-sm">Edit</a></td>
        <td><a href="/delete/${item._id}" class="btn btn-danger btn-sm" onclick="return confirm('Yakin hapus?')">Delete</a></td>
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
app.get("/delete/:id", async (req, res) => {
  await Laporan.findByIdAndDelete(req.params.id);
  res.redirect("/laporan");
});

// -------------------- Edit --------------------
app.get("/edit/:id", async (req, res) => {
  const item = await Laporan.findById(req.params.id).lean();
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
      <form action="/edit/${item._id}" method="post" class="card p-4 shadow-sm">
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
          <input name="waktuKerusakan" type="datetime-local" class="form-control">
        </div>
        <div class="mb-3">
          <label class="form-label">Perbaikan</label>
          <input name="perbaikan" value="${item.perbaikan}" class="form-control">
        </div>
        <div class="mb-3">
          <label class="form-label">Waktu Selesai</label>
          <input name="selesai" type="datetime-local" class="form-control">
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

app.post("/edit/:id", async (req, res) => {
  const { gardu, kerusakan, perbaikan, selesai, status, by, waktuKerusakan } = req.body;
  await Laporan.findByIdAndUpdate(req.params.id, {
    gardu,
    kerusakan,
    perbaikan,
    selesai: formatTanggal(selesai),
    status,
    by,
    waktuKerusakan: formatTanggal(waktuKerusakan),
  });
  res.redirect("/laporan");
});

// -------------------- Export Excel (OPTIMIZE) --------------------
app.get("/export/excel", async (req, res) => {
  const laporan = await Laporan.find().lean();
  const workbook = new ExcelJS.Workbook();
  const ws = workbook.addWorksheet("Laporan");

  ws.columns = [
    { header: "No", key: "no", width: 5 },
    { header: "Foto", key: "foto", width: 30 },
    { header: "Gardu", key: "gardu", width: 15 },
    { header: "Kerusakan", key: "kerusakan", width: 30 },
    { header: "Waktu Kerusakan", key: "waktuKerusakan", width: 20 },
    { header: "Perbaikan", key: "perbaikan", width: 30 },
    { header: "Waktu Selesai", key: "selesai", width: 20 },
    { header: "Status", key: "status", width: 15 },
    { header: "By", key: "by", width: 15 },
    { header: "Tanggal Upload", key: "tanggalUpload", width: 25 },
  ];

  for (let i = 0; i < laporan.length; i++) {
    const item = laporan[i];
    const row = ws.addRow({
      no: i + 1,
      foto: "",
      gardu: item.gardu,
      kerusakan: item.kerusakan,
      waktuKerusakan: item.waktuKerusakan,
      perbaikan: item.perbaikan,
      selesai: item.selesai,
      status: item.status,
      by: item.by,
      tanggalUpload: item.tanggalUpload,
    });

    try {
      const urlKecil = item.fotoURL.replace(
        "/upload/",
        "/upload/w_200,h_150,c_fill,q_auto/"
      );
      const response = await axios.get(urlKecil, { responseType: "arraybuffer" });
      const buffer = Buffer.from(response.data);

      const imageId = workbook.addImage({ buffer, extension: "jpeg" });
      ws.addImage(imageId, {
        tl: { col: 1, row: row.number - 1 },
        ext: { width: 100, height: 80 },
      });
    } catch (err) {
      console.log("⚠️ Gagal fetch image Excel:", err.message);
    }
  }

  res.setHeader(
    "Content-Type",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  );
  res.setHeader("Content-Disposition", "attachment; filename=laporan.xlsx");

  await workbook.xlsx.write(res);
  res.end();
});

// -------------------- Export Word (OPTIMIZE) --------------------
app.get("/export/word", async (req, res) => {
  const laporan = await Laporan.find().lean();

  const doc = new Document({
    sections: [
      {
        properties: {},
        children: await Promise.all(
          laporan.map(async (item, index) => {
            let children = [
              new Paragraph({ children: [new TextRun({ text: `No: ${index + 1}`, bold: true })] }),
            ];

            try {
              const urlKecil = item.fotoURL.replace(
                "/upload/",
                "/upload/w_300,h_200,c_fill,q_auto/"
              );
              const response = await axios.get(urlKecil, { responseType: "arraybuffer" });
              const buffer = Buffer.from(response.data);
              const image = Media.addImage(doc, buffer, 200, 150);
              children.push(new Paragraph(image));
            } catch (err) {
              console.log("⚠️ Gagal fetch image Word:", err.message);
            }

            children.push(
              new Paragraph(`Gardu: ${item.gardu}`),
              new Paragraph(`Kerusakan: ${item.kerusakan}`),
              new Paragraph(`Waktu Kerusakan: ${item.waktuKerusakan}`),
              new Paragraph(`Perbaikan: ${item.perbaikan}`),
              new Paragraph(`Waktu Selesai: ${item.selesai}`),
              new Paragraph(`Status: ${item.status}`),
              new Paragraph(`By: ${item.by}`),
              new Paragraph(`Tanggal Upload: ${item.tanggalUpload}`),
              new Paragraph("----------------------------")
            );

            return children;
          })
        ).then((arr) => arr.flat()),
      },
    ],
  });

  const buffer = await Packer.toBuffer(doc);
  res.setHeader("Content-Disposition", "attachment; filename=laporan.docx");
  res.send(buffer);
});

// -------------------- Run Server --------------------
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running at http://localhost:${PORT}`));

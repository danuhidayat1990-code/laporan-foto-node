const express = require("express");
const multer = require("multer");
const { CloudinaryStorage } = require("multer-storage-cloudinary");
const cloudinary = require("cloudinary").v2;
const ExcelJS = require("exceljs");
const { Document, Packer, Paragraph, TextRun, Media } = require("docx");
const mongoose = require("mongoose");
const axios = require("axios");
const sharp = require("sharp");
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

async function fetchAndCompressImage(url) {
  try {
    const response = await axios.get(url, { responseType: "arraybuffer" });
    let buffer = Buffer.from(response.data, "binary");
    // Kompres gambar biar gak gede
    buffer = await sharp(buffer).resize(300).jpeg({ quality: 70 }).toBuffer();
    return buffer;
  } catch (err) {
    console.log("⚠️ Gagal fetch image:", err.message);
    return null;
  }
}

// -------------------- Routes --------------------
app.get("/", (req, res) => {
  res.send(`
  <html>
  <head><title>Upload</title></head>
  <body>
    <form action="/upload" method="post" enctype="multipart/form-data">
      <input type="file" name="foto" required /><br/>
      <input type="text" name="gardu" placeholder="Gardu" required /><br/>
      <input type="text" name="kerusakan" placeholder="Kerusakan" required /><br/>
      <input type="datetime-local" name="waktuKerusakan" required /><br/>
      <input type="text" name="perbaikan" placeholder="Perbaikan" required /><br/>
      <input type="datetime-local" name="selesai" required /><br/>
      <select name="status">
        <option value="Selesai">Selesai</option>
        <option value="Proses">Proses</option>
        <option value="Pending">Pending</option>
      </select><br/>
      <input type="text" name="by" placeholder="By" required /><br/>
      <button type="submit">Upload</button>
    </form>
  </body>
  </html>
  `);
});

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

app.get("/laporan", async (req, res) => {
  const laporan = await Laporan.find().lean();
  let html = `<h2>Daftar Laporan</h2><table border="1" cellpadding="5"><tr>
  <th>No</th><th>Foto</th><th>Gardu</th><th>Kerusakan</th><th>Selesai</th></tr>`;
  laporan.forEach((item, i) => {
    html += `<tr><td>${i + 1}</td><td><img src="${item.fotoURL}" width="100"/></td>
    <td>${item.gardu}</td><td>${item.kerusakan}</td><td>${item.selesai}</td></tr>`;
  });
  html += `</table><a href="/export/excel">Export Excel</a> | <a href="/export/word">Export Word</a>`;
  res.send(html);
});

// -------------------- Export Excel --------------------
app.get("/export/excel", async (req, res) => {
  const laporan = await Laporan.find().lean();
  const workbook = new ExcelJS.Workbook();
  const ws = workbook.addWorksheet("Laporan");

  ws.columns = [
    { header: "No", key: "no", width: 5 },
    { header: "Foto", key: "foto", width: 20 },
    { header: "Gardu", key: "gardu", width: 15 },
    { header: "Kerusakan", key: "kerusakan", width: 30 },
    { header: "Selesai", key: "selesai", width: 20 },
  ];

  for (let i = 0; i < laporan.length; i++) {
    const item = laporan[i];
    const row = ws.addRow({ no: i + 1, gardu: item.gardu, kerusakan: item.kerusakan, selesai: item.selesai });
    const buffer = await fetchAndCompressImage(item.fotoURL);
    if (buffer) {
      const imageId = workbook.addImage({ buffer, extension: "jpeg" });
      ws.addImage(imageId, { tl: { col: 1, row: row.number - 1 }, ext: { width: 80, height: 60 } });
    }
  }

  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.setHeader("Content-Disposition", "attachment; filename=laporan.xlsx");
  await workbook.xlsx.write(res);
  res.end();
});

// -------------------- Export Word --------------------
app.get("/export/word", async (req, res) => {
  const laporan = await Laporan.find().lean();
  const doc = new Document({
    sections: [
      {
        children: await Promise.all(
          laporan.map(async (item, i) => {
            const children = [
              new Paragraph({ children: [new TextRun({ text: `No: ${i + 1}`, bold: true })] }),
            ];
            const buffer = await fetchAndCompressImage(item.fotoURL);
            if (buffer) {
              const image = Media.addImage(doc, buffer, 200, 150);
              children.push(new Paragraph(image));
            }
            children.push(new Paragraph(`Gardu: ${item.gardu}`));
            children.push(new Paragraph(`Kerusakan: ${item.kerusakan}`));
            children.push(new Paragraph(`Selesai: ${item.selesai}`));
            children.push(new Paragraph("-----------------------------"));
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
app.listen(PORT, () => console.log(`✅ Server running on http://localhost:${PORT}`))
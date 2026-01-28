const express = require("express");
const axios = require("axios");
const XLSX = require("xlsx");
const bodyParser = require("body-parser");

const app = express();
app.use(bodyParser.json({ limit: "10mb" }));
app.use(express.static(__dirname));

/* ================= UTIL ================= */

function cleanPhone(phone) {
  if (!phone) return "";
  return phone.replace(/[^0-9]/g, "").replace(/^62/, "0");
}

function cleanNumber(val) {
  if (!val) return "";
  return parseInt(val.toString().replace(/\./g, "")) || "";
}

function normalizeText(text) {
  if (!text) return "";
  return text.replace(/[,\.]{2,}/g, "").trim();
}

/* ====== Ambil Kode Pos via Internet ====== */
async function getKodePos(query) {
  try {
    const url = `https://kodepos.co.id/search?q=${encodeURIComponent(query)}`;
    const res = await axios.get(url, { timeout: 5000 });
    const match = res.data.match(/\b\d{5}\b/);
    return match ? match[0] : "";
  } catch {
    return "";
  }
}

/* ================= PROCESS ENGINE ================= */

app.post("/process", async (req, res) => {
  try {
    const rows = req.body.data.trim().split("\n");
    const output = [];

    for (let row of rows) {
      const cols = row.split("\t");

      let [
        cs, tgl, nama, hp, alamat,
        prov, kab, kec, kodepos,
        , , qty, harga, ongkir, pot
      ] = cols;

      cs = (cs || "").toUpperCase();
      hp = cleanPhone(hp);
      prov = (prov || "").toUpperCase();
      kab = (kab || "").toUpperCase();
      kec = (kec || "").toUpperCase();

      alamat = normalizeText(alamat);

      if (!kodepos) {
        kodepos = await getKodePos(`${alamat} ${kec}`);
      }

      alamat = `${alamat}, ${kodepos}, ${hp}`;

      output.push([
        cs,
        "",
        nama || "",
        hp,
        alamat,
        prov,
        kab,
        kec,
        kodepos || "",
        "",
        "",
        cleanNumber(qty),
        cleanNumber(harga),
        cleanNumber(ongkir),
        cleanNumber(pot)
      ]);
    }

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(output);
    XLSX.utils.book_append_sheet(wb, ws, "Data");

    const file = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });

    res.setHeader("Content-Disposition", "attachment; filename=hasil_kurir.xlsx");
    res.send(file);

  } catch (err) {
    console.error(err);
    res.status(500).send("Terjadi error saat memproses data.");
  }
});

/* ================= SERVER START ================= */

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log("Server jalan di port " + PORT));

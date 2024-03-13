const express = require("express");
const app = express();
const port = 3000;
const multer = require("multer");
const upload = multer({ dest: "uploads/" });
const xlsxPopulate = require("xlsx-populate");

app.listen(port, () =>
  console.log("> Server is up and running on port : " + port)
);
app.set("view engine", "ejs");

app.use(express.json());
app.use(express.urlencoded({ extended: true }));

app.get("/", (req, res) => {
  res.render("index.ejs");
});

app.post("/excel", upload.single("excel"), async (req, res) => {
  try {
    const workbook = await xlsxPopulate.fromFileAsync(req.file.path);
    const sheet = workbook.sheet("Hoja");
    const usedRange = sheet.usedRange();

    const lastRow = usedRange.endCell().rowNumber();

    const range = sheet.range(`A5:G${lastRow}`).value();

    const keys = range[0];

    const arrayDeObjetos = [];

    for (let i = 1; i < range.length; i++) {
      const objeto = {};
      for (let j = 0; j < keys.length; j++) {
        objeto[keys[j]] = range[i][j];
      }
      arrayDeObjetos.push(objeto);
    }

    console.table(arrayDeObjetos);
    
    res.send("Archivo procesado exitosamente.");
  } catch (error) {
    console.error("Error al cargar el archivo:", error);
    res.status(500).send("Ocurrió un error al procesar el archivo.");
  }
});
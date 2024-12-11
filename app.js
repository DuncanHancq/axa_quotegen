const express = require('express');
const { engine } = require('express-handlebars');
const multer = require('multer');
const bodyParser = require('body-parser');
const cors = require('cors');
const path = require('path');
const fs = require('fs');
const xlsx = require('xlsx');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const archiver = require('archiver');

const app = express();
const port = 3000;
app.use(cors());
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, 'public')));
app.use(express.static(path.join(__dirname, 'views')));
app.engine('handlebars', engine());
app.set('view engine', 'handlebars');

const timestamp = Date.now();
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    const uploadDir = path.join(__dirname, 'temp/uploads', `${timestamp}`);
    fs.mkdirSync(uploadDir, { recursive: true }); // Ensure the directory exists
    cb(null, uploadDir);
  },
  filename: (req, file, cb) => {
    cb(null, `${timestamp}-${file.originalname}`);
  }
});

const upload = multer({ storage });

app.post('/generate-quotes', upload.fields([{ name: 'xlsxFile' }, { name: 'docxFile' }]), async (req, res) => {
  try {
    const files = req.files;
    const body = req.body;

    const mapping = JSON.parse(body.mapping);
    const naming = JSON.parse(body.naming);

    // Load the XLSX file
    const workbook = xlsx.readFile(files.xlsxFile[0].path);
    const staticFileNaming = naming.staticFileNaming;
    const idNamingOptions = naming.idNamingOptions;
    const sortBy = naming.sortBy;
    const sheetName = naming.sheetSelect;

    const worksheet = workbook.Sheets[sheetName];
    const rows = xlsx.utils.sheet_to_json(worksheet);
    
    const docxContent = fs.readFileSync(files.docxFile[0].path, 'binary');

    rows.forEach((row, index) => {
      console.log(`Row ${index + 1}:`);
      const fullPath = path.join(__dirname, `temp/output/${timestamp}/${row[sortBy]}`, `${staticFileNaming}-${row[idNamingOptions]}.docx`);

      const zip = new PizZip(docxContent);
      const doc = new Docxtemplater(zip, {
        delimiters: { start: "<<", end: ">>" },
        linebreaks: true,
        paragraphLoop: true,
      });

      const data = {};
      for (const [docxField, excelColumn] of Object.entries(mapping)) {
        data[docxField] = row[excelColumn];
      }

      try {
        doc.render(data);
      } catch (error) {
        throw new Error(`Erreur lors du rendu du document : ${error.message}`);
      }

      const buffer = doc.getZip().generate({ type: 'nodebuffer' });

      if (!fs.existsSync(path.dirname(fullPath))) {
        fs.mkdirSync(path.dirname(fullPath), { recursive: true });
      }

      fs.writeFileSync(fullPath, buffer);
    });

    // Chemin du dossier à compresser
    const folderToZip = path.join(__dirname, `temp/output/${timestamp}`);
    const zipFilePath = path.join(__dirname, `temp/output/${timestamp}.zip`);

    // Compresser le dossier
    const output = fs.createWriteStream(zipFilePath);
    const archive = archiver('zip', { zlib: { level: 9 } });

    output.on('close', () => {
      console.log(`Archive créée avec succès : ${zipFilePath} (${archive.pointer()} octets)`);
      res.status(200).json({ message: 'Devis générés et archivés avec succès !', downloadCode: timestamp });
    });

    archive.on('error', (err) => {
      throw err;
    });

    archive.pipe(output);
    archive.directory(folderToZip, false);
    archive.finalize();

  } catch (error) {
    console.error('Erreur lors du traitement de la requête :', error);
    res.status(500).json({ error: 'Erreur serveur. Veuillez réessayer.' });
  }
});

app.get('/download-zip/:timestamp', (req, res) => {
  const timestamp = req.params.timestamp;
  const zipFilePath = path.join(__dirname, `temp/output/${timestamp}.zip`);

  if (fs.existsSync(zipFilePath)) {
    res.download(zipFilePath, `devis-${timestamp}.zip`, (err) => {
      if (err) {
        console.error('Erreur lors du téléchargement du fichier ZIP :', err);
        res.status(500).send('Erreur lors du téléchargement du fichier ZIP.');
      }
    });
  } else {
    res.status(404).send('Fichier ZIP non trouvé.');
  }
});

app.get('/', (req, res) => {
  res.render('home', { title: 'Quotegen' });
});

app.listen(port, () => {
  console.log(`Serveur démarré sur http://localhost:${port}`);
});


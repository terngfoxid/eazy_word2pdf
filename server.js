var express = require('express')
var multer = require('multer')
const path = require('path');
const fs = require('fs').promises;
var convertWordFiles = require('convert-multiple-files-ul');

var app = express()
app.use(multer({ dest: './uploads/' }).single('file'));

//ui route
app.get('/', function (req, res) {
    res.sendFile(__dirname + '/index.html')
})

app.get('/word2pdf_ui', function (req, res) {
    res.sendFile(__dirname + '/htdocs/word2pdf_ui.html')
})

//function route
app.post('/upload', function (req, res) {
    //req.file.originalname
    //req.file.mimetype
    const newPath = './uploads/' + req.file.originalname
    fs.rename(req.file.path, newPath, function (err) {
        if (err) {
            console.log('ERROR: ' + err);
            res.sendStatus(400)
        }
    });
    res.sendStatus(200)
})

app.post('/word2pdf', async function (req, res) {
    let fileName =[]
    switch (req.file.mimetype) {
        case "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
            fileName = req.file.originalname.split(".docx")
            break;
        case "application/msword":
            fileName = req.file.originalname.split(".doc")
            break;
        default: res.sendStatus(400);
    }
    if (fileName.length < 1) {
        res.sendStatus(400)
    }

    const newPath = './uploads/' + req.file.originalname
    await fs.rename(req.file.path, newPath, function (err) {
        if (err) {
            console.log('ERROR: ' + err);
            res.sendStatus(400)
        }
    });

    const sourceFilePath = newPath;
    //const outputFilePath = path.resolve('./uploads/' + fileName[0] + '.pdf');
    const outputFileDir = path.resolve('./uploads/');

    /*const docxBuf = await fs.readFile(sourceFilePath);
    let pdfBuf = await libre.convertAsync(docxBuf, '.pdf', undefined);
    await fs.writeFile(outputFilePath, pdfBuf);*/
    try {
        const pathOutput = await convertWordFiles.convertWordFiles(sourceFilePath, 'pdf', outputFileDir);
        res.sendFile(pathOutput);
    } catch (err) {
        console.log(err);
        res.sendStatus(500)
    }
    
    try {
        const files = await fs.readdir('./uploads');
    
        const deleteFilePromises = files.map(file =>
          fs.unlink(path.join('./uploads', file)),
        );
    
        await Promise.all(deleteFilePromises);
      } catch (err) {
        console.log(err);
      }
})

app.listen(8087, function () {
    console.log('App running on port 8087')
})
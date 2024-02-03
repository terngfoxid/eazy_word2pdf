var express = require('express')
var multer = require('multer')
const path = require('path');
const fs = require('fs').promises;
//var convertWordFiles = require('convert-multiple-files-ul');
var cors = require('cors')

const libre = require('libreoffice-convert');
libre.convertAsync = require('util').promisify(libre.convert);

var corsOptions = {
    origin: '*',
    optionsSuccessStatus: 200
}

var app = express()
app.use(cors(corsOptions));
app.use(multer({ dest: './uploads/' }).single('file'));

//ui route
app.get('/', function (req, res) {
    res.sendFile(__dirname + '/index.html')
})

app.get('/word2pdf_ui', function (req, res) {
    res.sendFile(__dirname + '/htdocs/word2pdf_ui.html')
})

app.get('/excel2pdf_ui', function (req, res) {
    res.sendFile(__dirname + '/htdocs/excel2pdf_ui.html')
})

//function route
app.post('/', async function (req, res) {
    res.sendStatus(200)
})

app.post('/word2pdf', async function (req, res) {
    let fileName = []
    switch (req.file.mimetype) {
        case "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
            fileName = req.file.originalname.split(".docx");
            break;
        case "application/msword":
            fileName = req.file.originalname.split(".doc");
            break;
        default: ;
    }
    if (fileName.length < 1) {
        fileName[0] = req.file.filename;
    }

    const newPath = './uploads/' + req.file.originalname
    await fs.rename(req.file.path, newPath, function (err) {
        if (err) {
            console.log('ERROR: ' + err);
            res.sendStatus(400)
        }
    });

    //const sourceFilePath = newPath;
    //const outputFilePath = path.resolve('./uploads/' + fileName[0] + '.pdf');
    //const outputFileDir = path.resolve('./uploads/');

    /*const docxBuf = await fs.readFile(sourceFilePath);
    let pdfBuf = await libre.convertAsync(docxBuf, '.pdf', undefined);
    await fs.writeFile(outputFilePath, pdfBuf);*/
    /*
    try {
        const pathOutput = await convertWordFiles.convertWordFiles(sourceFilePath, 'pdf', outputFileDir);
        res.sendFile(pathOutput)
    } catch (err) {
        console.log(err);
        res.sendStatus(500);
    }*/

    try {
        const ext = '.pdf'
        const inputPath = newPath;
        const outputPath = path.resolve('./uploads/' + fileName[0] + `${ext}`);

        // Read file
        const docxBuf = await fs.readFile(inputPath);

        // Convert it to pdf format with undefined filter (see Libreoffice docs about filter)
        let pdfBuf = await libre.convertAsync(docxBuf, ext, undefined);

        // Here in done you have pdf file which you can save or transfer in another stream
        await fs.writeFile(outputPath, pdfBuf);
        res.sendFile(outputPath);
    }catch (err){
        console.log(err);
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

app.post('/excel2pdf', async function (req, res) {
    let fileName = []
    switch (req.file.mimetype) {
        case "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
            fileName = req.file.originalname.split(".xlsx");
            break;
        case "application/vnd.ms-excel":
            fileName = req.file.originalname.split(".xls");
            break;
        default: ;
    }
    if (fileName.length < 1) {
        fileName[0] = req.file.filename;
    }

    const newPath = './uploads/' + req.file.originalname
    await fs.rename(req.file.path, newPath, function (err) {
        if (err) {
            console.log('ERROR: ' + err);
            res.sendStatus(400)
        }
    });

    //const sourceFilePath = newPath;
    //const outputFilePath = path.resolve('./uploads/' + fileName[0] + '.pdf');
    //const outputFileDir = path.resolve('./uploads/');

    /*const docxBuf = await fs.readFile(sourceFilePath);
    let pdfBuf = await libre.convertAsync(docxBuf, '.pdf', undefined);
    await fs.writeFile(outputFilePath, pdfBuf);*/
    /*
    try {
        const pathOutput = await convertWordFiles.convertWordFiles(sourceFilePath, 'pdf', outputFileDir);
        res.sendFile(pathOutput)
    } catch (err) {
        console.log(err);
        res.sendStatus(500);
    }*/

    try {
        const ext = '.pdf'
        const inputPath = newPath;
        const outputPath = path.resolve('./uploads/' + fileName[0] + `${ext}`);

        // Read file
        const docxBuf = await fs.readFile(inputPath);
        
        // Convert it to pdf format with undefined filter (see Libreoffice docs about filter)
        let pdfBuf = await libre.convertAsync(docxBuf, ext, undefined);

        // Here in done you have pdf file which you can save or transfer in another stream
        await fs.writeFile(outputPath, pdfBuf);
        res.sendFile(outputPath);
    }catch (err){
        console.log(err);
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

app.listen(8096, function () {
    console.log('App running on port 8096')
})
import express from "express";
import { dirname } from "path";
import { fileURLToPath } from "url";
import bodyParser from "body-parser";
import nodemailer from "nodemailer";
import qrcode from "qrcode";
import fs from "fs";
import { promisify } from "util";
import { createCanvas, loadImage } from "canvas";
import { v4 as uuidv4 } from 'uuid';
import XLSX from 'xlsx'; // Correct the import

const __dirname = dirname(fileURLToPath(import.meta.url));

const port = 3001;
const app = express();

app.use(bodyParser.urlencoded({ extended: true }));

const writeFile = promisify(fs.writeFile);

app.use(express.static('static')); // Serve static files from the 'static' directory

app.get("/", (req, res) => {
  res.sendFile(__dirname + "/static/index.html");
});

app.get("/ex-student", (req, res) => {
  res.sendFile(__dirname + "/static/form.html");
});

app.post("/submit", async (req, res) => {
  try {
    console.log(req.body);

    const name = req.body["Name"];
    const email = req.body["Email"];
    const no = req.body["phone"];
    const passOutYear = req.body["passoutyear"];

    // Generate a unique ID
    const uniqueId = uuidv4();

    // Generate a QR code as a data URL
    const qrCodeData = `ID: ${uniqueId}, Name: ${name}, Email: ${email}, Phone: ${no}, Year of Pass Out: ${passOutYear}`;
    const qrCodeDataURL = await qrcode.toDataURL(qrCodeData);

    // Create a composite image with the QR code and user details
    const templateImagePath = __dirname + '/static/template.png';
    const templateImage = await loadImage(templateImagePath);
    const canvas = createCanvas(templateImage.width, templateImage.height);
    const ctx = canvas.getContext('2d');

    // Draw template
    ctx.drawImage(templateImage, 0, 0);

    // Draw QR code (enlarged)
    const qrCode = await loadImage(qrCodeDataURL);
    const qrCodeSize = 1000; // Adjust this value to set the size of the QR code
    const qrX = (canvas.width - qrCodeSize) / 2;
    const qrY = (canvas.height / 2 - qrCodeSize) / 2 - 100; // Centered in the upper half
    ctx.drawImage(qrCode, qrX, qrY, qrCodeSize, qrCodeSize);

    // Add user details
    const textX = canvas.width / 2;
    const textY = canvas.height / 2 + 50; // Adjusted to be in the lower half
    ctx.textAlign = 'center';
    ctx.font = '70px Arial'; // Adjust font size as needed
    ctx.fillStyle = '#000';
    ctx.fillText(`Name: ${name}`, textX, textY);
    ctx.fillText(`Email: ${email}`, textX, textY + 60);
    ctx.fillText(`Phone: ${no}`, textX, textY + 120);
    ctx.fillText(`Year of Pass Out: ${passOutYear}`, textX, textY + 180);
    ctx.fillText(`ID: ${uniqueId}`, textX, textY + 240);

    // Save composite image to file
    const outputImagePath = `${__dirname}/static/composite_${email}.png`;
    const buffer = canvas.toBuffer('image/png');
    await writeFile(outputImagePath, buffer);

    // Save details to Excel sheet
    const excelFilePath = `${__dirname}/static/registrations.xlsx`;
    let workbook;
    let worksheet;

    // Check if the file exists
    if (fs.existsSync(excelFilePath)) {
      workbook = XLSX.readFile(excelFilePath);
      worksheet = workbook.Sheets[workbook.SheetNames[0]];
    } else {
      workbook = XLSX.utils.book_new();
      worksheet = XLSX.utils.aoa_to_sheet([["Name", "Email", "Phone", "Pass Out Year", "Unique ID", "Entered"]]);
      XLSX.utils.book_append_sheet(workbook, worksheet, "Registrations");
    }

    // Append the new row
    const newRow = [name, email, no, passOutYear, uniqueId, ""]; // "Entered" column left blank
    XLSX.utils.sheet_add_aoa(worksheet, [newRow], { origin: -1 });

    // Write to file
    XLSX.writeFile(workbook, excelFilePath);

    const htmlContent = `
      <html>
        <body>
          <p>Hello ${name},</p>
          <p>Thank you for submitting your information. We have received your details:</p>
          <ul>
            <li>Name: ${name}</li>
            <li>Email: ${email}</li>
            <li>Phone Number: ${no}</li>
            <li>Year of Pass Out: ${passOutYear}</li>
            <li>Unique ID: ${uniqueId}</li>
          </ul>
          <p>Attached is your unique QR code image.</p>
          <p>Best regards,<br>Your Company</p>
        </body>
      </html>`;

    // Configure the email transport using the default SMTP transport and a GMail account
    const transporter = nodemailer.createTransport({
      service: 'gmail',
      auth: {
        user: 'bhattacharjee.agniva.jobs@gmail.com', // Replace with your email address
        pass: 'dovz mfxv bfcy inpl'   // Replace with your email password or app-specific password
      }
    });

    // Email options with attachment
    const mailOptions = {
      from: 'bhattacharjee.agniva.jobs@gmail.com', // Replace with your email address
      to: email,
      subject: 'Your Event QR Code',
      html: htmlContent,
      attachments: [
        {
          filename: `composite_${email}.png`,
          path: outputImagePath,
          cid: 'unique@qr.code' // Same cid value as in the HTML img src
        }
      ]
    };

    // Send the email
    transporter.sendMail(mailOptions, (error, info) => {
      if (error) {
        return console.log(error);
      }
      console.log('Email sent: ' + info.response);

      // Clean up the composite image file
      fs.unlink(outputImagePath, (err) => {
        if (err) {
          console.error('Error deleting the composite image file:', err);
        }
      });
    });

    res.sendFile(__dirname + "/static/thenga.html");
  } catch (error) {
    console.error('Error processing the form submission:', error);
    res.status(500).send('Internal Server Error');
  }
});

app.listen(port, () => {
  console.log(`Server is running on port ${port}`);
});

import express from 'express';
import { createTransport } from 'nodemailer';
import { existsSync } from 'fs';
import dotenv from 'dotenv';
import XLSX from 'xlsx';
import bodyParser from 'body-parser';

dotenv.config();

const app = express();
const PORT = 3001;
const { utils, writeFile, readFile } = XLSX;

// Middleware
app.use(bodyParser.urlencoded({ extended: true }));
app.use(express.static('public'));

// Email configuration
const transporter = createTransport({
    service: 'gmail',
    auth: {
        user: process.env.EMAIL_USER,
        pass: process.env.EMAIL_PASS
    }
});

// Endpoint to handle form submission
app.post('/subscribe', (req, res) => {
    const email = req.body.email;
    const filePath = './subscribers.xlsx';

    // Check if file exists
    if (!existsSync(filePath)) {
        const wb = utils.book_new();
        const ws = utils.aoa_to_sheet([['Email']]);
        utils.book_append_sheet(wb, ws, 'Subscribers');
        writeFile(wb, filePath);
    }

    // Read existing data
    const workbook = readFile(filePath);
    const sheet = workbook.Sheets['Subscribers'];
    const data = utils.sheet_to_json(sheet, { header: 1 });

    // Check if email already exists
    const emailExists = data.some(row => row[0] === email);
    if (emailExists) {
        return res.send('Email already exists! We will notify you.');
    }

    // Add new email
    data.push([email]);

    // Write updated data back to the file
    const newSheet = utils.aoa_to_sheet(data);
    workbook.Sheets['Subscribers'] = newSheet;
    writeFile(workbook, filePath);

    // Send notification email
    const mailOptions = {
        from: process.env.EMAIL_USER,
        to: process.env.EMAIL_USER,
        subject: 'New Subscriber',
        text: `You have a new subscriber: ${email}`
    };

    transporter.sendMail(mailOptions, (error, info) => {
        if (error) {
            return console.log(error);
        }
        console.log('Email sent: ' + info.response);
    });

    res.send('Thank you for subscribing!');
});

app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});

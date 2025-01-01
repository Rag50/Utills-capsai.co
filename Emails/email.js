// emal getting code 




// const admin = require("firebase-admin");
// const axios = require("axios");

// // Initialize Firebase Admin SDK
// const serviceAccount = require("./caps-85254-firebase-adminsdk-31j3r-0edeb4bd98.json");

// admin.initializeApp({
//     credential: admin.credential.cert(serviceAccount),
// });

// // Web app URL from Google Apps Script
// const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbxmGkIDX6cYQXuo7y8Ha21VwpCgteH9kSNyfQrVLKe_JLrtOUs39u1xQiOrnbckxL5Z/exec";

// // Fetch all users and filter Google Auth users
// const fetchGoogleAuthUsers = async () => {
//     let nextPageToken;
//     const googleAuthEmails = [];

//     do {
//         const listUsersResult = await admin.auth().listUsers(1000, nextPageToken);
//         listUsersResult.users.forEach((user) => {
//             if (user.providerData.some((provider) => provider.providerId === "google.com")) {
//                 googleAuthEmails.push(user.email); // Collect emails
//             }
//         });
//         nextPageToken = listUsersResult.pageToken;
//     } while (nextPageToken);

//     return googleAuthEmails;
// };

// // Fetch existing emails from Google Sheets
// const fetchExistingEmails = async () => {
//     try {
//         const response = await axios.get(WEB_APP_URL);
//         return response.data; // Existing emails
//     } catch (error) {
//         console.error("Error fetching existing emails:", error);
//         return [];
//     }
// };

// // Send emails to Google Apps Script
// const sendToGoogleSheet = async (emails) => {
//     try {
//         const response = await axios.post(WEB_APP_URL, emails, {
//             headers: { "Content-Type": "application/json" },
//         });
//         console.log("Response from Google Apps Script:", response.data);
//     } catch (error) {
//         console.error("Error sending data to Google Sheet:", error);
//     }
// };

// // Main function
// const main = async () => {
//     try {
//         // Fetch emails from Firebase
//         const emails = await fetchGoogleAuthUsers();

//         // Fetch existing emails from Google Sheets
//         const existingEmails = await fetchExistingEmails();

//         // Filter out already existing emails
//         const newEmails = emails.filter(email => !existingEmails.includes(email));

//         // If there are new emails, send them to Google Sheets
//         if (newEmails.length > 0) {
//             const formattedEmails = newEmails.map(email => [email]); // Convert to 2D array for Sheets API
//             await sendToGoogleSheet(formattedEmails);
//         } else {
//             console.log("No new emails to add.");
//         }
//     } catch (error) {
//         console.error("Error:", error);
//     }
// };

// main();


// email sending code 

const nodemailer = require('nodemailer');
const xlsx = require('xlsx');
const path = require('path');


const EXCEL_FILE_PATH = path.join(__dirname, 'Emails.xlsx');


function getEmailsFromExcel() {
    const workbook = xlsx.readFile(EXCEL_FILE_PATH);
    const sheetName = workbook.SheetNames[0]; 
    const sheetData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 }); 

 
    const emails = sheetData.map((column) => column[0]).filter((email) => email && email.includes('@')); 
    return emails;
}


async function sendEmailsInBatches(emails, batchSize, delay) {
    try {
    
        const transporter = nodemailer.createTransport({
            service: 'gmail', 
            auth: {
                user: 'ai.editor@capsai.co', 
                pass: 'mkjg lqdt ubqa yiia', 
            },
        });

        for (let i = 0; i < emails.length; i += batchSize) {
            const batch = emails.slice(i, i + batchSize); 
            console.log(`Sending batch ${Math.ceil((i + 1) / batchSize)}...`);

            for (const email of batch) {
                const mailOptions = {
                    from: '"Capsai" <ai.editor@capsai.co>',
                    to: email,
                    subject: '🎉 Start 2025 with 10 Free Subtitle minutes on CapsAI!',
                    html: `
                        <!DOCTYPE html>
                        <html>
                        <head>
                            <style>
                                body {
                                    font-family: Arial, sans-serif;
                                    margin: 0;
                                    padding: 0;
                                    background-color: #f4f4f4;
                                }
                                .email-container {
                                    max-width: 600px;
                                    margin: 20px auto;
                                    background-color: #ffffff;
                                    border-radius: 10px;
                                    overflow: hidden;
                                    box-shadow: 0 2px 5px rgba(0, 0, 0, 0.1);
                                }
                                .image-container img {
                                    width: 100%;
                                    display: block;
                                }
                                .cta-button {
                                    display: block;
                                    width: 90%;
                                    max-width: 400px;
                                    margin: 20px auto;
                                    padding: 15px;
                                    background-color: #000000;
                                    color: #ffffff;
                                    text-align: center;
                                    text-decoration: none;
                                    border-radius: 5px;
                                    font-size: 18px;
                                    font-weight: bold;
                                }
                                .cta-button:hover {
                                    background-color: #444444;
                                }
                            </style>
                        </head>
                        <body>
                            <div class="email-container">
                                <div class="image-container">
                                    <a href="https://capsai.co/" target="_blank">
                                        <img src="https://capsaistore.blob.core.windows.net/capsaiassets/2025-email.png" alt="Celebrate 2025 with Free AI Subtitles">
                                    </a>
                                </div>
                            </div>
                        </body>
                        </html>
                    `,
                };

                // Send email
                await transporter.sendMail(mailOptions);
                console.log(`Email sent to ${email}`);
            }

            console.log(`Batch ${Math.ceil((i + 1) / batchSize)} completed. Waiting ${delay / 1000} seconds before the next batch...`);
            await new Promise((resolve) => setTimeout(resolve, delay)); // Delay between batches
        }

        console.log('All emails sent successfully.');
    } catch (error) {
        console.error('Error sending emails:', error);
    }
}


(async () => {
    const emails = getEmailsFromExcel();
    console.log(`Fetched ${emails.length} emails from the Excel file.`);

    const batchSize = 30; 
    const delay = 60000; 

    await sendEmailsInBatches(emails, batchSize, delay);
})();
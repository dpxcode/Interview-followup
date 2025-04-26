const XLSX = require('xlsx');
const nodemailer = require('nodemailer');
require('dotenv').config();

// Create a transporter using SMTP
const transporter = nodemailer.createTransport({
    service: 'gmail',
    auth: {
        user: process.env.EMAIL_USER,
        pass: process.env.EMAIL_PASS
    }
});

// Function to read Excel file
function readExcelFile(filePath) {
    try {
        const workbook = XLSX.readFile(filePath);
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        return XLSX.utils.sheet_to_json(worksheet);
    } catch (error) {
        console.error('Error reading Excel file:', error);
        return [];
    }
}

// Function to update Excel file with new data
function updateExcelFile(filePath, data) {
    try {
        const workbook = XLSX.readFile(filePath);
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        
        // Convert data to worksheet format
        const newWorksheet = XLSX.utils.json_to_sheet(data);
        
        // Replace the existing worksheet
        workbook.Sheets[sheetName] = newWorksheet;
        
        // Write the updated workbook
        XLSX.writeFile(workbook, filePath);
        return true;
    } catch (error) {
        console.error('Error updating Excel file:', error);
        return false;
    }
}

// Function to read follow-up tracking file
function readFollowUpTracking(filePath) {
    try {
        const workbook = XLSX.readFile(filePath);
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        return XLSX.utils.sheet_to_json(worksheet);
    } catch (error) {
        console.error('Error reading follow-up tracking file:', error);
        return [];
    }
}

// Function to update follow-up tracking file
function updateFollowUpTracking(filePath, data) {
    try {
        const workbook = XLSX.readFile(filePath);
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        
        // Convert data to worksheet format
        const newWorksheet = XLSX.utils.json_to_sheet(data);
        
        // Replace the existing worksheet
        workbook.Sheets[sheetName] = newWorksheet;
        
        // Write the updated workbook
        XLSX.writeFile(workbook, filePath);
        return true;
    } catch (error) {
        console.error('Error updating follow-up tracking file:', error);
        return false;
    }
}

// Function to update follow-up status
function updateFollowUpStatus(email, status, notes = '') {
    const trackingFile = 'followup_tracking.xlsx';
    const trackingData = readFollowUpTracking(trackingFile);
    
    const contactIndex = trackingData.findIndex(contact => contact.email === email);
    if (contactIndex === -1) {
        console.log(`Contact ${email} not found in tracking file`);
        return false;
    }
    
    // Update the status
    trackingData[contactIndex].responseStatus = status;
    trackingData[contactIndex].responseDate = new Date().toISOString();
    if (notes) {
        trackingData[contactIndex].notes = notes;
    }
    
    return updateFollowUpTracking(trackingFile, trackingData);
}

// Function to check if it's time to send a follow-up
function shouldSendFollowUp(lastSentDate, includeResume = false) {
    if (!lastSentDate) return true; // First time sending
    
    const lastSent = new Date(lastSentDate);
    const now = new Date();
    const daysSinceLastSent = Math.floor((now - lastSent) / (1000 * 60 * 60 * 24));
    
    if (includeResume) {
        // Send with resume every 7 days
        return daysSinceLastSent >= 7;
    } else {
        // Send regular follow-up every 2 days
        return daysSinceLastSent >= 2;
    }
}

// Function to send email
async function sendEmail(recipient) {
    const mailOptions = {
        from: process.env.EMAIL_USER,
        to: recipient.email,
        subject: 'Application for Full Stack Developer (Node.js, React.js, Python) - Durga Parshad',
        html: `
            <div style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
                <p>Dear ${recipient.name || 'HR Professional'},</p>
                
                <p>I hope this email finds you well. I am writing to express my strong interest in the Full Stack Developer position at ${recipient.company || 'your organization'}.</p>
                
                <p>As a passionate Full Stack Developer with comprehensive expertise in modern web technologies and cloud platforms, I bring a unique combination of technical skills that I believe would make me a valuable addition to your team. Here's what I offer:</p>
                
                <ul>
                    <li><strong>Full Stack Development:</strong>
                        <ul>
                            <li>Backend expertise in Node.js, Python, and RESTful APIs</li>
                            <li>Frontend mastery with React.js and modern UI/UX practices</li>
                            <li>Willingness to learn Java Spring Boot for enterprise applications</li>
                        </ul>
                    </li>
                    <li><strong>Cloud & DevOps:</strong>
                        <ul>
                            <li>Strong knowledge of AWS and GCP cloud platforms</li>
                            <li>Experience with cloud architecture and deployment strategies</li>
                        </ul>
                    </li>
                    <li><strong>AI & Integration:</strong>
                        <ul>
                            <li>Hands-on experience with OpenAI integrations and API implementations</li>
                            <li>Building intelligent features using AI/ML services</li>
                        </ul>
                    </li>
                    <li><strong>Modern Development:</strong>
                        <ul>
                            <li>Proficient in version control, CI/CD pipelines, and agile methodologies</li>
                            <li>Strong focus on code quality, testing, and best practices</li>
                        </ul>
                    </li>
                </ul>
                
                <p>I am particularly drawn to ${recipient.company || 'your company'} because of your commitment to innovation and your reputation in the industry. I am confident that my technical skills, cloud expertise, and experience with cutting-edge technologies align perfectly with your team's needs.</p>
                
                <p>I have attached my resume for your review. I would welcome the opportunity to discuss how my skills and experience could contribute to your team's success. I am available for an interview at your convenience.</p>
                
                <p>Thank you for considering my application. I look forward to the possibility of joining your team.</p>
                
                <p>Best regards,<br>
                Durga Parshad</p>
            </div>
        `,
        attachments: [
            {
                filename: 'Durga_Parshad_Resume.pdf',
                path: './Durga_Parshad.pdf'
            }
        ]
    };

    try {
        await transporter.sendMail(mailOptions);
        console.log(`Email sent successfully to ${recipient.email}`);
        return true;
    } catch (error) {
        console.error(`Error sending email to ${recipient.email}:`, error);
        return false;
    }
}

// Modified sendFollowUpEmail function to update tracking
async function sendFollowUpEmail(recipient, includeResume) {
    const trackingFile = 'followup_tracking.xlsx';
    const trackingData = readFollowUpTracking(trackingFile);
    
    // Find or create tracking entry
    let trackingEntry = trackingData.find(contact => contact.email === recipient.email);
    if (!trackingEntry) {
        trackingEntry = {
            email: recipient.email,
            name: recipient.name,
            company: recipient.company,
            lastSentDate: new Date().toISOString(),
            followUpCount: 0,
            responseStatus: 'Pending',
            responseDate: null,
            notes: ''
        };
        trackingData.push(trackingEntry);
    }
    
    // Update follow-up count
    trackingEntry.followUpCount = (trackingEntry.followUpCount || 0) + 1;
    trackingEntry.lastSentDate = new Date().toISOString();
    
    // Update tracking file
    updateFollowUpTracking(trackingFile, trackingData);
    
    const followUpOptions = {
        from: process.env.EMAIL_USER,
        to: recipient.email,
        subject: includeResume 
            ? 'Follow-up with Resume: Application for Full Stack Developer (Node.js, React.js, Python) - Durga Parshad'
            : 'Follow-up: Application for Full Stack Developer (Node.js, React.js, Python) - Durga Parshad',
        html: `
            <div style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
                <p>Dear ${recipient.name || 'HR Professional'},</p>
                
                <p>I hope this email finds you well. I recently applied for the Full Stack Developer position at ${recipient.company || 'your organization'} and wanted to follow up on my application.</p>
                
                <p>I am still very interested in the position and would appreciate any feedback you could provide regarding my application. If there are any additional materials or information you need from me, please let me know.</p>
                
                <p>I understand that the hiring process can take time, and I appreciate your consideration of my application. I would welcome the opportunity to discuss how my skills in Node.js, React.js, Python, and my willingness to learn Java Spring Boot could benefit your team.</p>
                
                ${includeResume ? '<p>I have attached my resume again for your convenience.</p>' : ''}
                
                <p>Thank you for your time and consideration.</p>
                
                <p>Best regards,<br>
                Durga Parshad</p>
            </div>
        `
    };
    
    // Add resume attachment if requested
    if (includeResume) {
        followUpOptions.attachments = [
            {
                filename: 'Durga_Parshad_Resume.pdf',
                path: './Durga_Parshad.pdf'
            }
        ];
    }

    try {
        await transporter.sendMail(followUpOptions);
        console.log(`Follow-up email ${includeResume ? 'with resume' : ''} sent successfully to ${recipient.email}`);
        return true;
    } catch (error) {
        console.error(`Error sending follow-up email to ${recipient.email}:`, error);
        return false;
    }
}

// Function to process follow-up emails for all contacts
async function processFollowUpEmails() {
    const hrList = readExcelFile('hr_contacts.xlsx');
    
    if (hrList.length === 0) {
        console.log('No contacts found in the Excel file');
        return;
    }

    console.log(`Found ${hrList.length} contacts. Checking for follow-up emails...`);
    
    const updatedContacts = [];
    let emailsSent = 0;
    const now = new Date();
    
    for (const contact of hrList) {
        if (!contact.email) continue;
        
        const lastSentDate = contact.lastSentDate ? new Date(contact.lastSentDate) : null;
        const daysSinceLastSent = lastSentDate ? Math.floor((now - lastSentDate) / (1000 * 60 * 60 * 24)) : null;
        
        // Determine if we should send a follow-up
        let shouldSend = false;
        let includeResume = false;
        
        if (!lastSentDate) {
            // First time sending
            shouldSend = true;
            includeResume = true;
        } else if (daysSinceLastSent >= 7) {
            // Time for a resume follow-up
            shouldSend = true;
            includeResume = true;
        }
        
        if (shouldSend) {
            const success = await sendFollowUpEmail(contact, includeResume);
            if (success) {
                emailsSent++;
                // Update the last sent date
                contact.lastSentDate = now.toISOString();
            }
        } else if (lastSentDate) {
            console.log(`Skipping ${contact.email} - Last email was sent ${daysSinceLastSent} days ago`);
        }
        
        updatedContacts.push(contact);
        
        // Add a delay between emails to avoid rate limiting
        await new Promise(resolve => setTimeout(resolve, 2000));
    }
    
    // Update the Excel file with new last sent dates
    if (emailsSent > 0) {
        updateExcelFile('hr_contacts.xlsx', updatedContacts);
        console.log(`Updated Excel file with new last sent dates. Sent ${emailsSent} follow-up emails.`);
    } else {
        console.log('No follow-up emails needed - all contacts have been emailed within the last 7 days.');
    }
}

// Main function to process the Excel file and send initial emails
async function processEmails() {
    const hrList = readExcelFile('hr_contacts.xlsx');
    
    if (hrList.length === 0) {
        console.log('No contacts found in the Excel file');
        return;
    }

    console.log(`Found ${hrList.length} contacts. Starting to send emails...`);
    
    const updatedContacts = [];
    let emailsSent = 0;
    const now = new Date();
    
    for (const contact of hrList) {
        if (!contact.email) continue;
        
        // Check if 7 days have passed since last email
        const lastSentDate = contact.lastSentDate ? new Date(contact.lastSentDate) : null;
        const daysSinceLastSent = lastSentDate ? Math.floor((now - lastSentDate) / (1000 * 60 * 60 * 24)) : null;
        
        // Only send if no previous email or 7 days have passed
        if (!lastSentDate || (daysSinceLastSent && daysSinceLastSent >= 7)) {
            const success = await sendEmail(contact);
            if (success) {
                emailsSent++;
                // Set the initial last sent date
                contact.lastSentDate = now.toISOString();
            }
        } else {
            console.log(`Skipping ${contact.email} - Last email was sent ${daysSinceLastSent} days ago`);
        }
        
        updatedContacts.push(contact);
        
        // Add a delay between emails to avoid rate limiting
        await new Promise(resolve => setTimeout(resolve, 2000));
    }
    
    // Update the Excel file with new last sent dates
    if (emailsSent > 0) {
        updateExcelFile('hr_contacts.xlsx', updatedContacts);
        console.log(`Updated Excel file with new last sent dates. Sent ${emailsSent} initial emails.`);
    } else {
        console.log('No initial emails needed - all contacts have been emailed within the last 7 days.');
    }
}

// Function to initialize tracking for existing contacts
async function initializeTracking() {
    const hrContacts = readExcelFile('hr_contacts.xlsx');
    const trackingFile = 'followup_tracking.xlsx';
    let trackingData = [];
    
    for (const contact of hrContacts) {
        trackingData.push({
            email: contact.email,
            name: contact.name,
            company: contact.company,
            lastSentDate: contact.lastSentDate || null,
            followUpCount: 0,
            responseStatus: 'Pending',
            responseDate: null,
            notes: ''
        });
    }
    
    return updateFollowUpTracking(trackingFile, trackingData);
}

// Function to check response status
function checkResponseStatus() {
    const trackingFile = 'followup_tracking.xlsx';
    const trackingData = readFollowUpTracking(trackingFile);
    
    console.log('\nCurrent Application Status:');
    console.log('------------------------');
    
    trackingData.forEach(contact => {
        console.log(`\nCompany: ${contact.company}`);
        console.log(`Contact: ${contact.name} (${contact.email})`);
        console.log(`Status: ${contact.responseStatus}`);
        console.log(`Follow-ups sent: ${contact.followUpCount}`);
        if (contact.responseDate) {
            console.log(`Last response: ${new Date(contact.responseDate).toLocaleDateString()}`);
        }
        if (contact.notes) {
            console.log(`Notes: ${contact.notes}`);
        }
    });
}

// Function to send no-response follow-up email
async function sendNoResponseFollowUp(recipient) {
    const mailOptions = {
        from: process.env.EMAIL_USER,
        to: recipient.email,
        subject: 'Following Up: Application for Full Stack Developer (Node.js, React.js, Python) - Durga Parshad',
        html: `
            <div style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
                <p>Dear ${recipient.name || 'HR Professional'},</p>
                
                <p>I hope this email finds you well. I recently submitted my application for the Full Stack Developer position at ${recipient.company || 'your organization'} and I wanted to follow up as I haven't received a response yet.</p>
                
                <p>I understand that you may be reviewing many applications, and I wanted to reiterate my strong interest in the position. As a Full Stack Developer with expertise in Node.js, React.js, and Python, I believe I would be a valuable addition to your team. I am also eager to learn Java Spring Boot to contribute to your enterprise applications.</p>
                
                <p>I would greatly appreciate any feedback regarding my application or information about the next steps in your hiring process. I have attached my resume again for your convenience.</p>
                
                <p>Thank you for your time and consideration. I look forward to your response.</p>
                
                <p>Best regards,<br>
                Durga Parshad</p>
            </div>
        `,
        attachments: [
            {
                filename: 'Durga_Parshad_Resume.pdf',
                path: './Durga_Parshad.pdf'
            }
        ]
    };

    try {
        await transporter.sendMail(mailOptions);
        console.log(`No-response follow-up email sent successfully to ${recipient.email}`);
        return true;
    } catch (error) {
        console.error(`Error sending no-response follow-up email to ${recipient.email}:`, error);
        return false;
    }
}

// Function to process no-response follow-ups
async function processNoResponseFollowUps() {
    const hrList = readExcelFile('hr_contacts.xlsx');
    const trackingData = readFollowUpTracking('followup_tracking.xlsx');
    
    if (hrList.length === 0) {
        console.log('No contacts found in the Excel file');
        return;
    }

    console.log('Checking for contacts who haven\'t responded...');
    
    const updatedContacts = [];
    let emailsSent = 0;
    
    for (const contact of hrList) {
        if (!contact.email) continue;
        
        // Find tracking entry
        const trackingEntry = trackingData.find(entry => entry.email === contact.email);
        
        // Send follow-up if:
        // 1. Initial email was sent (has lastSentDate)
        // 2. No response received (responseStatus is 'Pending' or undefined)
        // 3. No follow-up sent in the last 5 days
        if (contact.lastSentDate && 
            (!trackingEntry || trackingEntry.responseStatus === 'Pending') &&
            (!trackingEntry?.lastSentDate || 
             (new Date() - new Date(trackingEntry.lastSentDate)) / (1000 * 60 * 60 * 24) >= 5)) {
            
            const success = await sendNoResponseFollowUp(contact);
            if (success) {
                emailsSent++;
                // Update tracking
                if (trackingEntry) {
                    trackingEntry.lastSentDate = new Date().toISOString();
                    trackingEntry.followUpCount = (trackingEntry.followUpCount || 0) + 1;
                } else {
                    trackingData.push({
                        email: contact.email,
                        name: contact.name,
                        company: contact.company,
                        lastSentDate: new Date().toISOString(),
                        followUpCount: 1,
                        responseStatus: 'Pending',
                        responseDate: null,
                        notes: 'No-response follow-up sent'
                    });
                }
            }
        }
        
        updatedContacts.push(contact);
        
        // Add a delay between emails
        await new Promise(resolve => setTimeout(resolve, 2000));
    }
    
    // Update tracking file
    if (emailsSent > 0) {
        updateFollowUpTracking('followup_tracking.xlsx', trackingData);
        console.log(`Sent ${emailsSent} no-response follow-up emails.`);
    } else {
        console.log('No no-response follow-up emails needed at this time.');
    }
}

// Check command line arguments to determine which function to run
const args = process.argv.slice(1);
if (args.includes('--follow-up')) {
    // Run the follow-up email function
    processFollowUpEmails().catch(console.error);
} else if (args.includes('--no-response')) {
    // Run the no-response follow-up function
    processNoResponseFollowUps().catch(console.error);
} else if (args.includes('--check-status')) {
    // Check response status
    checkResponseStatus();
} else if (args.includes('--update-status')) {
    // Update status for a specific contact
    const email = args[args.indexOf('--email') + 1];
    const status = args[args.indexOf('--status') + 1];
    const notes = args[args.indexOf('--notes') + 1] || '';
    
    if (!email || !status) {
        console.log('Please provide email and status using --email and --status flags');
        process.exit(1);
    }
    
    if (updateFollowUpStatus(email, status, notes)) {
        console.log('Status updated successfully');
    } else {
        console.log('Failed to update status');
    }
} else if (args.includes('--init-tracking')) {
    // Initialize tracking for existing contacts
    if (initializeTracking()) {
        console.log('Successfully initialized tracking for existing contacts');
    } else {
        console.log('Failed to initialize tracking');
    }
} else {
    // Run the main application
    processEmails().catch(console.error);
} 
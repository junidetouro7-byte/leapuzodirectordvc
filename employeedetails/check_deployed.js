const fs = require('fs');
const content = fs.readFileSync('C:\\Users\\JUNIDE CHRIS\\.gemini\\antigravity\\brain\\4c893571-0834-45ad-9d65-466345a48e93\\.system_generated\\steps\\412\\content.md', 'utf8');

// Check if our styles exist
console.log('Includes copyright flex override:', content.includes('.framer-lrmfcq .framer-1znye5'));
console.log('Includes Ph No clean text:', content.includes('Ph No : +91-9384815318'));
console.log('Includes leading space:', content.includes(' Ph No : +91-9384815318'));

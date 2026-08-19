const fs = require('fs');
const content = fs.readFileSync('framer_index.html', 'utf8');

// Find occurrences of framer-YCag4 in the HTML
const regex = /framer-YCag4/g;
let count = 0;
while (regex.exec(content) !== null) {
    count++;
}
console.log('Occurrences of framer-YCag4 in framer_index.html:', count);

// Search for the element that has this class in the HTML markup
const elementRegex = /<[^>]*?framer-YCag4[^>]*?>/gi;
let match;
while ((match = elementRegex.exec(content)) !== null) {
    console.log(match[0]);
}

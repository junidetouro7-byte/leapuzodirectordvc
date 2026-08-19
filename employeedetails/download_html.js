const https = require('https');

https.get('https://leapuzo-employee-detailsfinal.vercel.app/', (res) => {
    let data = '';
    res.on('data', (chunk) => {
        data += chunk;
    });
    res.on('end', () => {
        console.log('HTML Length:', data.length);
        const styleStart = data.indexOf('/* Force exact left alignment for copyright');
        if (styleStart !== -1) {
            console.log('Found custom styles! Snippet:');
            console.log(data.substring(styleStart, styleStart + 500));
        } else {
            console.log('Custom styles NOT found in deployed HTML!');
            // Let's print the custom style block head part
            const customStart = data.indexOf('<!-- Start of Custom Styles');
            if (customStart !== -1) {
                console.log('Found custom block start! Snippet:');
                console.log(data.substring(customStart, customStart + 1000));
            } else {
                console.log('Custom block start NOT found in deployed HTML!');
            }
        }
    });
}).on('error', (err) => {
    console.error('Error fetching live site:', err.message);
});

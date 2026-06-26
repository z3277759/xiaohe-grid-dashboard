const os = require('os');
const nets = os.networkInterfaces();
let found = false;
for (const name of Object.keys(nets)) {
    for (const net of nets[name]) {
        if (net.family === 'IPv4' && !net.internal) {
            console.log('=============================');
            console.log('Share this URL with colleagues:');
            console.log('http://' + net.address + ':3000');
            console.log('=============================');
            found = true;
        }
    }
}
if (!found) {
    console.log('No LAN IP found. Check network connection.');
}

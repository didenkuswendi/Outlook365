const outlook365 = require('./lib/outlook365.js');

// wajib menjalankan chrome debug "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --remote-debugging-address=127.0.0.1

(async () => {        
    const email = await outlook365(
        "didenkuswendi@akunemail.com",
        "didenkuswendi@akunemail.com",
        "ini judul",
        "<strong>Halo, ini email pertama</strong>"
    )
    if(email){await email.disconnect()}
})();

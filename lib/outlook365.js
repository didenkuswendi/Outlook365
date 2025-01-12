const puppeteer = require('puppeteer-core');

async function outlook365(to,cc,body) {
    // Hubungkan Puppeteer ke browser aktif
    const browserURL = 'http://localhost:9222'; // Alamat debugging
    const browser = await puppeteer.connect({ browserURL });

    // Ambil tab yang terbuka
    const pages = await browser.pages();
    const page = pages.find((p) => p.url().includes('outlook.office.com'));

    if (!page) {
        console.error('Halaman Outlook Web tidak ditemukan. Pastikan Anda sudah membuka Outlook Web di browser.');
        return;
    }


    // Dapatkan sesi Chrome DevTools Protocol (CDP)
    const session = await page.target().createCDPSession();    
    const { windowId, bounds } = await session.send('Browser.getWindowForTarget');
    const windowState = bounds.windowState;

    // Cek apakah browser minimize
    if (windowState != 'minimized') {                
        await session.send('Browser.setWindowBounds', {
            windowId,
            bounds: { windowState: 'minimized' },
        });
    }

    if (bounds.windowState === 'minimized' || bounds.windowState === 'fullscreen') {
    // Ubah jendela ke mode normal terlebih dahulu
    await session.send('Browser.setWindowBounds', {
        windowId,
        bounds: { windowState: 'normal' },
    });
    }    

    // Bawa tab ke depan
    await page.bringToFront();
    console.log('Tab Outlook Web telah diaktifkan.');

    // Maksimalkan jendela setelah mode normal
    await session.send('Browser.setWindowBounds', {
    windowId,
    bounds: { windowState: 'maximized' },
    });

    // Tunggu elemen tombol "Email baru"
    await page.waitForSelector('[aria-label="Email baru"]', { visible: true });

    // Klik tombol "Baru pesan"
    await page.click('[aria-label="Email baru"]');

    // Tunggu form pesan terbuka
    await page.waitForSelector('[aria-label="Kepada"]', { visible: true });

    // Isi form email
    await page.type('[aria-label="Kepada"]', to) // Alamat penerima
    await page.type('[aria-label="Cc"]', cc); // Alamat Cc
    await page.type('input[aria-label="Tambahkan subjek"]', 'Test Email via Browser Aktif'); // Subjek email
    await page.type('div[aria-label="Isi pesan, tekan Alt+F10 untuk keluar"]', body); // Isi email

    // Klik tombol kirim
    await page.click('button[aria-label="Kirim"]');

    // Tunggu beberapa saat untuk memastikan email terkirim
    await new Promise((resolve) => setTimeout(resolve, 3000));
    // minimize kembali windows
    await session.send('Browser.setWindowBounds', {
        windowId,
        bounds: { windowState: 'minimized' },
    });
    console.log('Email berhasil dikirim!');    
    return browser; 
}

module.exports = outlook365;
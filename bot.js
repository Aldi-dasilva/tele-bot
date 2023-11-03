const TelegramBot = require('node-telegram-bot-api');
const ExcelJS = require('exceljs');

// Token API bot Telegram
const TOKEN = 'masukan_token_bot-anda';  // Gantilah dengan token bot Telegram Anda
const EXCEL_FILE_PATH = './database/datausers.xlsx';  // Gantilah dengan path file Excel yang sesuai

// Inisialisasi bot
const bot = new TelegramBot(TOKEN, { polling: true });

// Objek untuk menyimpan state pencarian pelanggan
const cariStates = {};

// Fungsi start
bot.onText(/\/start/, (msg) => {
  const chatId = msg.chat.id;
  bot.sendMessage(chatId, 'Selamat Datang! Ketik /Cari untuk mencari detail pelanggan.');
});

// Fungsi search_customer
bot.onText(/\/Cari/, (msg) => {
  const chatId = msg.chat.id;
  cariStates[chatId] = 'waiting_id'; // Menyimpan state pencarian

  bot.sendMessage(chatId, 'Masukkan ID Pelanggan:');
});

// Menangani pesan yang diinput oleh pengguna
bot.on('message', (msg) => {
  console.log('Pesan yang diterima:', msg.text);
  const chatId = msg.chat.id;
  const text = msg.text;

  if (cariStates[chatId] === 'waiting_id') {
    console.log('State: waiting_id');
    cariStates[chatId] = null; // Reset state

    // Konek ke Excel
    const workbook = new ExcelJS.Workbook();

    workbook.xlsx.readFile(EXCEL_FILE_PATH)
      .then(() => {
        const worksheet = workbook.getWorksheet(1);
        console.log('Membaca file Excel');

        // Cari data pelanggan berdasarkan ID
        let customer_data = {};
        worksheet.eachRow((row, rowNumber) => {
          if (rowNumber === 1) return; // Skip header row
        
          const rowData = row.values.filter((value) => value !== null);
          if (rowData[0] === text) {
            console.log('Data ditemukan:', rowData);
            customer_data = rowData;
          }
        });
        
        console.log('Data pelanggan:', customer_data);
        // menampilkan data excel pada bot
        if (customer_data) {
          console.log('Mengisi customer_data:', customer_data);
          console.log('Data pelanggan dari Excel:', customer_data);
          const [id_pelanggan, nama_lengkap, email, alamat] = customer_data;
          const message = `Detail Pelanggan:\nID Pelanggan: ${id_pelanggan}\nNama Lengkap: ${nama_lengkap}\nEmail: ${email}\nAlamat: ${alamat}`;
          bot.sendMessage(chatId, message);
        } else {
          bot.sendMessage(chatId, 'ID Pelanggan tidak ditemukan.');
        }        
      })
      .catch((error) => {
        console.error('Kesalahan saat mengakses file Excel:', error);
        bot.sendMessage(chatId, 'Terjadi kesalahan saat mengakses file Excel.');
      });
  }
});

// Tangani pesan selain perintah /start dan /search_customer
bot.onText(/^(?!\/start|\/Cari)/, (msg) => {
  const chatId = msg.chat.id;
  bot.sendMessage(chatId, 'Ketik /Cari untuk mencari detail pelanggan.');
});

(function () {
  // 1. INJECT PUSTAKA PPTXGENJS SECARA DINAMIS
  const pptxScript = document.createElement('script');
  pptxScript.src = "https://cdn.jsdelivr.net/gh/gitbrent/pptxgenjs@3.12.0/dist/pptxgen.bundle.js";
  document.head.appendChild(pptxScript);

  // 2. INJECT PUSTAKA GOOGLE DRIVE PICKER / API
  const gdriveScript = document.createElement('script');
  gdriveScript.src = "https://apis.google.com/js/api.js";
  document.head.appendChild(gdriveScript);

  // 3. FUNGSI INJEKSI MODUL EXTRA KE FORM
  function injectExtraFeatures() {
    // Hindari duplikasi jika elemen sudah pernah terpasang
    if (document.getElementById('btnGenPPT')) return;

    const form = document.getElementById('makalahForm');
    const btnSubmit = document.getElementById('btnSubmit');

    if (!form || !btnSubmit) return;

    const sectionCard = document.createElement('div');
    sectionCard.className = 'section-card';
    sectionCard.innerHTML = `
      <div class="section-title">🚀 Fitur Ekstra (PowerPoint & Drive)</div>
      <div class="row">
        <div class="form-group">
          <button type="button" id="btnGenPPT" class="btn-autofill" style="width:100%; background:#2b579a; border-color:#1e395b; box-shadow:0 4px 0 #1e395b;">
            📊 Generate Presentation (.pptx)
          </button>
        </div>
        <div class="form-group">
          <button type="button" id="btnSaveDrive" class="btn-history-toggle" style="width:100%; background:#0f9d58; border-color:#0b6b3a; box-shadow:0 4px 0 #0b6b3a;">
            ☁️ Simpan ke Google Drive
          </button>
        </div>
      </div>
    `;

    // Sisipkan sebelum tombol submit
    form.insertBefore(sectionCard, btnSubmit);

    // EVENT HANDLER: GENERATE PPTX CLIENT-SIDE
    document.getElementById('btnGenPPT').addEventListener('click', () => {
      const judul = document.getElementById('judulMakalah').value.trim();
      const nama = document.getElementById('nama').value.trim() || 'Mahasiswa';
      const kampus = document.getElementById('universitas').value.trim() || 'Kampus';

      if (!judul) {
        alert('Silakan isi "Judul Makalah" terlebih dahulu sebelum membuat slide presentasi.');
        document.getElementById('judulMakalah').focus();
        return;
      }

      if (typeof PptxGenJS === 'undefined') {
        alert('Library PowerPoint sedang dimuat, silakan coba 3 detik lagi.');
        return;
      }

      let pptx = new PptxGenJS();

      // Slide 1: Judul
      let slide1 = pptx.addSlide();
      slide1.addText(judul.toUpperCase(), { x: 0.5, y: 1.5, w: '90%', fontSize: 22, bold: true, color: '363636', align: 'center' });
      slide1.addText(`Disusun Oleh:\n${nama}\n\n${kampus}`, { x: 0.5, y: 3.5, w: '90%', fontSize: 14, color: '7F7F7F', align: 'center' });

      // Slide 2: Pendahuluan
      let slide2 = pptx.addSlide();
      slide2.addText('BAB I: PENDAHULUAN', { x: 0.5, y: 0.5, fontSize: 20, bold: true, color: '5B8C1D' });
      slide2.addText([
        { text: 'Latar belakang topik & urgency pembahasan.' },
        { text: 'Rumusan masalah penelitian.' },
        { text: 'Tujuan & batasan pembahasan makalah.' }
      ], { x: 0.5, y: 1.5, fontSize: 15, bullet: true });

      // Slide 3: Pembahasan
      let slide3 = pptx.addSlide();
      slide3.addText('BAB II: PEMBAHASAN', { x: 0.5, y: 0.5, fontSize: 20, bold: true, color: '5B8C1D' });
      slide3.addText([
        { text: 'Tinjauan teori dan landasan akademis.' },
        { text: 'Analisis dan temuan utama materi.' },
        { text: 'Implikasi serta penerapan dalam studi kasus.' }
      ], { x: 0.5, y: 1.5, fontSize: 15, bullet: true });

      // Slide 4: Kesimpulan
      let slide4 = pptx.addSlide();
      slide4.addText('BAB III: KESIMPULAN & SARAN', { x: 0.5, y: 0.5, fontSize: 20, bold: true, color: '5B8C1D' });
      slide4.addText('Ringkasan hasil pembahasan serta saran akademis untuk pengembangan ke depan.', { x: 0.5, y: 1.5, fontSize: 15 });

      const cleanTitle = judul.replace(/[^a-zA-Z0-9]/g, '_').substring(0, 20);
      pptx.writeFile({ fileName: `Slide_${cleanTitle}.pptx` });
    });

    // EVENT HANDLER: SIMPAN KE GOOGLE DRIVE
    document.getElementById('btnSaveDrive').addEventListener('click', () => {
      const judul = document.getElementById('judulMakalah').value.trim() || 'Makalah';
      
      const driveUrl = `https://drive.google.com/share?url=${encodeURIComponent(window.location.href)}&title=${encodeURIComponent(judul)}`;
      window.open(driveUrl, '_blank', 'width=600,height=500');
    });
  }

  // EKSEKUSI INJEKSI SETELAH DOM SIAP
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', injectExtraFeatures);
  } else {
    injectExtraFeatures();
  }
})();
// Ganti dengan Nomor WhatsApp Admin Anda (Gunakan kode negara 62, contoh: 628123456789)
const ADMIN_WA_NUMBER = "6282214630200"; 

function setupPaymentModal() {
  const modal = document.getElementById('paymentModal');
  const btnClose = document.getElementById('btnCloseModal');
  const btnCloseX = document.getElementById('btnCloseModalX');
  const btnConfirm = document.getElementById('btnConfirmPayment');

  if (!modal) return;

  // Pasang listener pada semua tombol Beli / Paket Pembayaran
  document.querySelectorAll('.btn-buy, .btn-beli').forEach(button => {
    button.addEventListener('click', (e) => {
      const card = e.target.closest('.card') || e.target.parentElement;
      const packageName = card ? (card.querySelector('h3, .package-name')?.innerText || 'Paket AI') : 'Paket AI';
      const price = e.target.getAttribute('data-price') || card?.querySelector('.price')?.innerText || 'Rp 10.000';

      document.getElementById('paymentPackage').innerText = packageName;
      document.getElementById('paymentAmount').innerText = price;
      modal.style.display = 'flex';
    });
  });

  // Fungsi Tutup Modal
  const closeModal = () => { modal.style.display = 'none'; };
  if (btnClose) btnClose.addEventListener('click', closeModal);
  if (btnCloseX) btnCloseX.addEventListener('click', closeModal);

  // Kirim Bukti via WhatsApp
  if (btnConfirm) {
    btnConfirm.addEventListener('click', () => {
      const paket = document.getElementById('paymentPackage').innerText;
      const nominal = document.getElementById('paymentAmount').innerText;
      
      const pesan = `Halo Admin, saya ingin konfirmasi pembayaran QRIS:\n\n` +
                    `📦 *Paket:* ${paket}\n` +
                    `💰 *Nominal:* ${nominal}\n\n` +
                    `Berikut saya lampirkan foto/screenshot bukti transfernya. Mohon diproses ya, terima kasih!`;

      const urlWA = `https://wa.me/${ADMIN_WA_NUMBER}?text=${encodeURIComponent(pesan)}`;
      window.open(urlWA, '_blank');
      closeModal();
    });
  }
}

// Inisialisasi setelah DOM dimuat
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', setupPaymentModal);
} else {
  setupPaymentModal();
}
// Fungsi Memeriksa Kuota dan Mengunci Form jika Belum Bayar/Habis Kuota
function checkUserAccess(userData) {
  const quotaElem = document.getElementById('quotaRemaining');
  const expiryElem = document.getElementById('expiryRemaining');
  const badgeElem = document.getElementById('userPackageBadge');
  const btnSubmit = document.getElementById('btnSubmit');

  // 1. Tampilkan Data di UI
  quotaElem.innerText = `${userData.kuotaMakalah} Makalah`;
  expiryElem.innerText = hitungSisaHari(userData.tanggalKadaluarsa);
  badgeElem.innerText = userData.paket;

  // 2. Cek Apakah Kuota Habis atau Masa Aktif Kadaluarsa
  const isExpired = userData.tanggalKadaluarsa && new Date(userData.tanggalKadaluarsa) < new Date();
  const isQuotaEmpty = userData.kuotaMakalah <= 0;

  if (isQuotaEmpty || isExpired || !userData.sudahBayar) {
    // Kunci tombol submit jika tidak punya akses
    if (btnSubmit) {
      btnSubmit.disabled = true;
      btnSubmit.style.background = '#ccc';
      btnSubmit.style.cursor = 'not-allowed';
      btnSubmit.innerText = '🔒 Beli Paket untuk Generate Makalah';
      
      // Jika diklik, tampilkan modal QRIS pembayaran
      btnSubmit.onclick = (e) => {
        e.preventDefault();
        alert('Kuota Anda sudah habis atau masa aktif paket telah berakhir. Silakan beli paket terlebih dahulu.');
        document.getElementById('paymentModal').style.display = 'flex';
      };
    }
  } else {
    // Izinkan pembuatan makalah
    if (btnSubmit) {
      btnSubmit.disabled = false;
      btnSubmit.innerText = '🚀 Generate Makalah Sekarang';
    }
  }
}
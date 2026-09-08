(function () {
  // 1. INJECT PUSTAKA PPTXGENJS SECARA DINAMIS
  const pptxScript = document.createElement('script');
  pptxScript.src = "https://cdn.jsdelivr.net/gh/gitbrent/pptxgenjs@3.12.0/dist/pptxgen.bundle.js";
  document.head.appendChild(pptxScript);

  // 2. INJECT PUSTAKA GOOGLE DRIVE PICKER / API
  const gdriveScript = document.createElement('script');
  gdriveScript.src = "https://apis.google.com/js/api.js";
  document.head.appendChild(gdriveScript);

  // 3. TAMBAHKAN TOMBOL BARU KE CONTAINER FORM KETIKA DOM SIAP
  window.addEventListener('DOMContentLoaded', () => {
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

    // Sisipkan sebelum tombol submit tanpa mengubah struktur HTML lama
    const form = document.getElementById('makalahForm');
    const btnSubmit = document.getElementById('btnSubmit');
    if (form && btnSubmit) {
      form.insertBefore(sectionCard, btnSubmit);
    }

    // EVENT HANDLER: GENERATE PPTX CLIENT-SIDE
    document.getElementById('btnGenPPT').addEventListener('click', () => {
      const judul = document.getElementById('judulMakalah').value || 'Makalah AI';
      const nama = document.getElementById('nama').value || 'Mahasiswa';
      const kampus = document.getElementById('universitas').value || 'Kampus';

      if (typeof PptxGenJS === 'undefined') {
        alert('Library PowerPoint sedang dimuat, silakan coba 3 detik lagi.');
        return;
      }

      let pptx = new PptxGenJS();

      // Slide 1: Judul
      let slide1 = pptx.addSlide();
      slide1.addText(judul, { x: 0.5, y: 1.5, w: '90%', fontSize: 24, bold: true, color: '363636', align: 'center' });
      slide1.addText(`Oleh: ${nama}\n${kampus}`, { x: 0.5, y: 3.5, w: '90%', fontSize: 16, color: '7F7F7F', align: 'center' });

      // Slide 2: Pendahuluan
      let slide2 = pptx.addSlide();
      slide2.addText('BAB I: PENDAHULUAN', { x: 0.5, y: 0.5, fontSize: 20, bold: true, color: '5B8C1D' });
      slide2.addText([
        { text: 'Latar Belakang masalah yang diangkat dalam penelitian.' },
        { text: 'Rumusan Masalah & Tujuan Pembahasan.' }
      ], { x: 0.5, y: 1.5, fontSize: 15, bullet: true });

      // Slide 3: Kesimpulan
      let slide3 = pptx.addSlide();
      slide3.addText('BAB III: KESIMPULAN', { x: 0.5, y: 0.5, fontSize: 20, bold: true, color: '5B8C1D' });
      slide3.addText('Ringkasan hasil pembahasan serta saran akademis untuk pengembangan ke depan.', { x: 0.5, y: 1.5, fontSize: 15 });

      pptx.writeFile({ fileName: `Slide_${judul.substring(0, 20)}.pptx` });
    });

    // EVENT HANDLER: SIMPAN KE GOOGLE DRIVE
    document.getElementById('btnSaveDrive').addEventListener('click', () => {
      const judul = document.getElementById('judulMakalah').value || 'Makalah';
      
      // Simulasi/Pemicu Upload Drive menggunakan Web Intent/Google Drive Save API
      const dummyContent = `Judul: ${judul}\nTanggal: ${new Date().toLocaleDateString()}`;
      const blob = new Blob([dummyContent], { type: 'text/plain' });
      const fileUrl = URL.createObjectURL(blob);

      // Membuka antarmuka penyimpanan Google Drive langsung
      const driveUrl = `https://drive.google.com/share?url=${encodeURIComponent(window.location.href)}&title=${encodeURIComponent(judul)}`;
      window.open(driveUrl, '_blank', 'width=600,height=500');
    });
  });
})();
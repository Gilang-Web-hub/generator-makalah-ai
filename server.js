import express from 'express';
import multer from 'multer';
import { GoogleGenAI } from '@google/genai';
import * as fs from 'fs';
import path from 'path';
import libre from 'libreoffice-convert';
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  AlignmentType,
  PageBreak,
  Footer,
  PageNumber,
  NumberFormat,
  convertMillimetersToTwip,
  convertInchesToTwip,
  ImageRun,
  HeadingLevel,
  TableOfContents
} from 'docx';

libre.convertAsync = function (docxBuf, format, filter) {
  return new Promise((resolve, reject) => {
    libre.convert(docxBuf, format, filter, (err, done) => {
      if (err) return reject(err);
      resolve(done);
    });
  });
};

const app = express();
const port = process.env.PORT || 3000;

// Gunakan Environment Variable jika ada, atau fallback ke key default
const apiKey = process.env.GEMINI_API_KEY || "AQ.Ab8RN6KEa_lUaASGKpTDU9t2jwqy_5DdAo3miZ2DLBDVp0El3A";
const ai = new GoogleGenAI({ apiKey });
const upload = multer({ dest: 'uploads/' });

app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(express.static('public'));

function cmToTwip(cm) {
  let num = parseFloat(cm);
  if (isNaN(num) || num < 0) {
    num = 3;
  }
  return convertMillimetersToTwip(num * 10);
}

function parseMarkdownItalicToRuns(textStr, fontName, fontSize) {
  const text = String(textStr || "").trim();
  const parts = text.split(/(\*[^*]+\*)/g);

  return parts.filter(p => p.length > 0).map(part => {
    if (part.startsWith('*') && part.endsWith('*')) {
      return new TextRun({
        text: part.slice(1, -1),
        italic: true,
        font: fontName,
        size: fontSize
      });
    } else {
      return new TextRun({
        text: part,
        font: fontName,
        size: fontSize
      });
    }
  });
}

function formatProdi(prodiStr) {
  const str = String(prodiStr || "").trim();
  if (!str) return "";
  if (str.toUpperCase().startsWith("PROGRAM STUDI") || str.toUpperCase().startsWith("PRODI")) {
    return str.toUpperCase();
  }
  return `PROGRAM STUDI ${str.toUpperCase()}`;
}

function formatFakultas(fakultasStr) {
  const str = String(fakultasStr || "").trim();
  if (!str) return "";
  if (str.toUpperCase().startsWith("FAKULTAS")) {
    return str.toUpperCase();
  }
  return `FAKULTAS ${str.toUpperCase()}`;
}

async function generateAiContentWithFallback(prompt, systemInstruction) {
  const candidateModels = ['gemini-2.5-flash', 'gemini-2.5-pro'];

  for (const modelName of candidateModels) {
    console.log(`[AI Request] Memilih model: ${modelName}...`);
    for (let attempt = 1; attempt <= 2; attempt++) {
      try {
        const response = await ai.models.generateContent({
          model: modelName,
          contents: prompt,
          config: {
            systemInstruction: systemInstruction,
            temperature: 0.2,
            responseMimeType: "application/json"
          },
        });
        console.log(`[AI Success] Berhasil diproses menggunakan ${modelName}`);
        return response;
      } catch (err) {
        console.warn(`[AI Warning] ${modelName} percobaan ${attempt} gagal: ${err.message}`);
        if (attempt < 2) {
          await new Promise(resolve => setTimeout(resolve, 2000));
        }
      }
    }
    console.log(`[AI Fallback] Beralih ke model berikutnya...`);
  }

  throw new Error("Seluruh server AI sedang sibuk. Silakan tunggu beberapa detik dan tekan tombol generate kembali.");
}

async function generateDocxBuffer(input, uploadedFilePath) {
  const systemInstruction = `
  Kamu adalah pakar akademis profesional di Indonesia. Tugasmu adalah menghasilkan konten makalah akademik yang sangat mendalam, rapi, dan presisi dalam bentuk JSON terstruktur.
  
  Aturan Ketat:
  1. Hasilkan TEPAT ${input.jumlahSubBab || 5} Sub-bab pada BAB II Pembahasan.
  2. Pada BAB II Pembahasan, SETIAP SUB-BAB WAJIB MEMILIKI MINIMAL 3 HINGGA 4 PARAGRAF PANJANG.
  3. Hasilkan TEPAT ${input.jumlahSubBab || 5} Rumusan Masalah pada BAB I.
  4. Hasilkan TEPAT ${input.jumlahSubBab || 5} Tujuan pada BAB I.
  5. SETIAP PARAGRAF pada BAB II WAJIB menyertakan sitasi ilmiah APA Style (Contoh: (Pratama & Wijaya, 2023)).
  6. Hasilkan TEPAT ${input.jumlahReferensi || 5} Daftar Pustaka bereputasi (2021-2026).
  7. PENTING UNTUK DAFTAR PUSTAKA: Bungkus HANYA judul buku atau judul artikel jurnal menggunakan tanda bintang tunggal (*contoh judul*).
     Contoh: Chen, L., Zhang, Y., & Wang, X. (2022). *Artificial intelligence in cardiovascular disease risk prediction*. Journal of Medical Systems, 46(3), 15-27.
  
  Format Output JSON Murni:
  {
    "kataPengantar": ["Paragraf 1...", "Paragraf 2..."],
    "latarBelakang": ["Paragraf 1...", "Paragraf 2..."],
    "rumusanMasalah": ["Pertanyaan 1...", "Pertanyaan 2..."],
    "tujuan": ["Tujuan 1...", "Tujuan 2..."],
    "pembahasan": [
      {
        "nomor": "2.1",
        "judul": "Judul Sub-bab 1",
        "paragraf": [
          "Paragraf 1...",
          "Paragraf 2...",
          "Paragraf 3..."
        ]
      }
    ],
    "kesimpulan": ["Paragraf 1..."],
    "saran": ["Paragraf 1..."],
    "daftarPustaka": [
      "Chen, L., Zhang, Y., & Wang, X. (2022). *Artificial intelligence in cardiovascular disease risk prediction*. Journal of Medical Systems, 46(3), 15-27."
    ]
  }
  `;

  const prompt = `Buatkan konten makalah komprehensif dan mendalam untuk:
  - Judul: ${input.judulMakalah}
  - Mata Kuliah: ${input.mataKuliah}
  - Jumlah Sub-bab BAB II: ${input.jumlahSubBab}
  - Jumlah Daftar Pustaka: ${input.jumlahReferensi}`;

  const response = await generateAiContentWithFallback(prompt, systemInstruction);
  const content = JSON.parse(response.text);

  const fontName = "Times New Roman";
  const fontSize = 24; 
  const lineSpacing = 360;

  const makeHeading1 = (text) => new Paragraph({
    alignment: AlignmentType.CENTER,
    heading: HeadingLevel.HEADING_1,
    spacing: { line: lineSpacing, after: 60 },
    children: [new TextRun({ text: String(text || ""), bold: true, font: fontName, size: fontSize, color: "000000" })]
  });

  const makeHeading2 = (text) => new Paragraph({
    alignment: AlignmentType.LEFT,
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 180, after: 60, line: lineSpacing },
    children: [new TextRun({ text: String(text || ""), bold: true, font: fontName, size: fontSize, color: "000000" })]
  });

  const makeParagraph = (text) => new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    indent: { firstLine: convertInchesToTwip(0.25) },
    spacing: { line: lineSpacing, after: 120 },
    children: [new TextRun({ text: String(text || ""), font: fontName, size: fontSize })]
  });

  const coverChildren = [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { line: lineSpacing, after: 60 },
      children: [new TextRun({ text: "MAKALAH", bold: true, font: fontName, size: fontSize })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { line: lineSpacing, after: 120 },
      children: [new TextRun({ text: String(input.judulMakalah || "").toUpperCase(), bold: true, font: fontName, size: fontSize })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { line: lineSpacing, after: 240 },
      children: [new TextRun({ text: `Disusun Untuk Memenuhi Tugas Mata Kuliah ${input.mataKuliah || ""}`, font: fontName, size: fontSize })]
    })
  ];

  let hasValidLogo = false;
  if (uploadedFilePath && fs.existsSync(uploadedFilePath)) {
    try {
      const logoBuffer = fs.readFileSync(uploadedFilePath);
      if (logoBuffer.length > 0) {
        coverChildren.push(
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 180, after: 240 },
            children: [
              new ImageRun({
                data: logoBuffer,
                transformation: { width: 120, height: 120 }
              })
            ]
          })
        );
        hasValidLogo = true;
      }
    } catch (imgErr) {
      console.error("Gagal menyisipkan logo:", imgErr);
    }
  }

  if (!hasValidLogo) {
    coverChildren.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 240, after: 240 },
        children: [new TextRun({ text: "[ LOGO UNIVERSITAS ]", bold: true, font: fontName, size: fontSize })]
      })
    );
  }

  coverChildren.push(
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { line: lineSpacing }, children: [new TextRun({ text: "Disusun Oleh:", bold: true, font: fontName, size: fontSize })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { line: lineSpacing }, children: [new TextRun({ text: `NAMA : ${String(input.nama || "").toUpperCase()}`, bold: true, font: fontName, size: fontSize })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { line: lineSpacing }, children: [new TextRun({ text: `NIM : ${input.nim || ""}`, bold: true, font: fontName, size: fontSize })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { line: lineSpacing }, children: [new TextRun({ text: `KELAS : ${input.kelas || ""}`, bold: true, font: fontName, size: fontSize })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { line: lineSpacing, after: 180 }, children: [new TextRun({ text: `SEMESTER : ${String(input.semester || "").toUpperCase()}`, bold: true, font: fontName, size: fontSize })] }),
    
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { line: lineSpacing }, children: [new TextRun({ text: "Dosen Pengampu:", bold: true, font: fontName, size: fontSize })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { line: lineSpacing, after: 240 }, children: [new TextRun({ text: String(input.namaDosen || "").toUpperCase(), bold: true, font: fontName, size: fontSize })] }),
    
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { line: lineSpacing }, children: [new TextRun({ text: formatProdi(input.prodi), bold: true, font: fontName, size: fontSize })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { line: lineSpacing }, children: [new TextRun({ text: formatFakultas(input.fakultas), bold: true, font: fontName, size: fontSize })] }),
    
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { line: lineSpacing }, children: [new TextRun({ text: String(input.universitas || "").toUpperCase(), bold: true, font: fontName, size: fontSize })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { line: lineSpacing }, children: [new TextRun({ text: String(input.kota || "").toUpperCase(), bold: true, font: fontName, size: fontSize })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { line: lineSpacing }, children: [new TextRun({ text: String(input.tahun || ""), bold: true, font: fontName, size: fontSize })] })
  );

  const kataPengantarParas = Array.isArray(content.kataPengantar) ? content.kataPengantar : [content.kataPengantar || ""];
  
  const pelengkapChildren = [
    makeHeading1("KATA PENGANTAR"),
    new Paragraph({ spacing: { after: 180 } }),
    ...kataPengantarParas.map(p => makeParagraph(p)),
    
    new Paragraph({
      alignment: AlignmentType.RIGHT,
      spacing: { before: 240, line: lineSpacing },
      children: [new TextRun({ text: `${input.kota || ""}, ${input.tanggal || ""}`, font: fontName, size: fontSize })]
    }),

    new Paragraph({ spacing: { line: lineSpacing }, children: [new TextRun({ text: "" })] }),
    new Paragraph({ spacing: { line: lineSpacing }, children: [new TextRun({ text: "" })] }),
    new Paragraph({ spacing: { line: lineSpacing }, children: [new TextRun({ text: "" })] }),

    new Paragraph({
      alignment: AlignmentType.RIGHT,
      spacing: { line: lineSpacing },
      children: [new TextRun({ text: input.nama || "", font: fontName, size: fontSize })]
    }),
    new Paragraph({ children: [new PageBreak()] }),

    makeHeading1("DAFTAR ISI"),
    new Paragraph({ spacing: { after: 180 } }),
    
    new TableOfContents("Daftar Isi", {
      hyperlink: true,
      headingStyleRange: "1-2"
    })
  ];

  const latarBelakang = Array.isArray(content.latarBelakang) ? content.latarBelakang : [content.latarBelakang || ""];
  const rumusanMasalah = Array.isArray(content.rumusanMasalah) ? content.rumusanMasalah : [content.rumusanMasalah || ""];
  const tujuan = Array.isArray(content.tujuan) ? content.tujuan : [content.tujuan || ""];
  const kesimpulan = Array.isArray(content.kesimpulan) ? content.kesimpulan : [content.kesimpulan || ""];
  const saran = Array.isArray(content.saran) ? content.saran : [content.saran || ""];
  const pembahasan = Array.isArray(content.pembahasan) ? content.pembahasan : [];
  const daftarPustaka = Array.isArray(content.daftarPustaka) ? content.daftarPustaka : [];

  const isiUtamaChildren = [
    makeHeading1("BAB I\nPENDAHULUAN"),
    makeHeading2("1.1 Latar Belakang"),
    ...latarBelakang.map(p => makeParagraph(p)),
    
    makeHeading2("1.2 Rumusan Masalah"),
    ...rumusanMasalah.map((p, i) => new Paragraph({
      alignment: AlignmentType.JUSTIFIED,
      spacing: { line: lineSpacing, after: 60 },
      children: [new TextRun({ text: `${i + 1}. ${p}`, font: fontName, size: fontSize })]
    })),
    
    makeHeading2("1.3 Tujuan"),
    ...tujuan.map((p, i) => new Paragraph({
      alignment: AlignmentType.JUSTIFIED,
      spacing: { line: lineSpacing, after: 60 },
      children: [new TextRun({ text: `${i + 1}. ${p}`, font: fontName, size: fontSize })]
    })),
    
    new Paragraph({ children: [new PageBreak()] }),
    makeHeading1("BAB II\nPEMBAHASAN"),
    ...pembahasan.flatMap(sub => {
      const subParas = Array.isArray(sub.paragraf) ? sub.paragraf : [sub.paragraf || ""];
      return [
        makeHeading2(`${sub.nomor || ""} ${sub.judul || ""}`),
        ...subParas.map(p => makeParagraph(p))
      ];
    }),
    
    new Paragraph({ children: [new PageBreak()] }),
    makeHeading1("BAB III\nKESIMPULAN DAN SARAN"),
    makeHeading2("3.1 Kesimpulan"),
    ...kesimpulan.map(p => makeParagraph(p)),
    
    makeHeading2("3.2 Saran"),
    ...saran.map(p => makeParagraph(p)),
    
    new Paragraph({ children: [new PageBreak()] }),
    makeHeading1("DAFTAR PUSTAKA"),
    
    ...daftarPustaka.map(itemStr => {
      const runs = parseMarkdownItalicToRuns(itemStr, fontName, fontSize);

      return new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        indent: { left: convertInchesToTwip(0.5), hanging: convertInchesToTwip(0.5) },
        spacing: { line: lineSpacing, after: 120 },
        children: runs
      });
    })
  ];

  const doc = new Document({
    features: {
      updateFields: true
    },
    sections: [
      {
        properties: {
          page: {
            margin: {
              top: cmToTwip(input.marginTop),
              bottom: cmToTwip(input.marginBottom),
              left: cmToTwip(input.marginLeft),
              right: cmToTwip(input.marginRight)
            }
          }
        },
        children: coverChildren
      },
      {
        properties: {
          page: {
            margin: {
              top: cmToTwip(input.marginTop),
              bottom: cmToTwip(input.marginBottom),
              left: cmToTwip(input.marginLeft),
              right: cmToTwip(input.marginRight)
            },
            pageNumberStart: 2,
            pageNumberFormatType: NumberFormat.LOWER_ROMAN
          }
        },
        footers: {
          default: new Footer({
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: [PageNumber.CURRENT], font: fontName, size: fontSize })] })]
          })
        },
        children: pelengkapChildren
      },
      {
        properties: {
          page: {
            margin: {
              top: cmToTwip(input.marginTop),
              bottom: cmToTwip(input.marginBottom),
              left: cmToTwip(input.marginLeft),
              right: cmToTwip(input.marginRight)
            },
            pageNumberStart: 1,
            pageNumberFormatType: NumberFormat.DECIMAL
          }
        },
        footers: {
          default: new Footer({
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: [PageNumber.CURRENT], font: fontName, size: fontSize })] })]
          })
        },
        children: isiUtamaChildren
      }
    ]
  });

  return await Packer.toBuffer(doc);
}

app.post('/api/generate-makalah', upload.single('logo'), async (req, res) => {
  let uploadedFilePath = null;

  try {
    const input = req.body;
    const formatOutput = (input.formatOutput || 'docx').toLowerCase();

    if (req.file) {
      uploadedFilePath = req.file.path;
    }

    console.log(`\n[Web Request] Memproses makalah: "${input.judulMakalah}" | Format: ${formatOutput.toUpperCase()}`);

    const bufferDocx = await generateDocxBuffer(input, uploadedFilePath);
    const fileNameBase = `Makalah_${(input.nama || 'Pengguna').replace(/\s+/g, '_')}`;

    if (formatOutput === 'pdf') {
      const pdfBuffer = await libre.convertAsync(bufferDocx, '.pdf', undefined);

      res.setHeader('Content-Type', 'application/pdf');
      res.setHeader('Content-Disposition', `attachment; filename="${fileNameBase}.pdf"`);
      res.send(pdfBuffer);

    } else {
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
      res.setHeader('Content-Disposition', `attachment; filename="${fileNameBase}.docx"`);
      res.send(bufferDocx);
    }

  } catch (error) {
    console.error("Error backend:", error);
    res.status(500).json({ error: "Gagal membuat makalah", details: error.message });
  } finally {
    // Pastikan file sementara terhapus setelah pengolahan selesai
    if (uploadedFilePath && fs.existsSync(uploadedFilePath)) {
      try {
        fs.unlinkSync(uploadedFilePath);
      } catch (e) {
        console.error("Gagal menghapus file temporer:", e);
      }
    }
  }
});

app.listen(port, () => {
  console.log(`\n Server Web Makalah Aktif & Siap!`);
  console.log(` Akses di browser: http://localhost:${port}`);
});
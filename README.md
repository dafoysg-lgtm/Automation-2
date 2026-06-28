# Invoice Reconciliation Automation 📄
> Semi-automated invoice reconciliation for PPh 23 contractor payments
> — cut per-file processing time from ~70s to ~20s across 300+ documents.

## Problem
Divisi pajak PT Indocement Tunggal Prakarsa Tbk memproses 
*Contractor Labor Payment Request* secara manual:

- Setiap pasangan file (PDF invoice + Excel) harus dibuka dan dicocokkan satu per satu
- Komponen yang diverifikasi: Plant/Division, Job Description, dan Subtotal
- Jika tidak sinkron → potensi denda pajak dan cancel
- Setelah verifikasi: Nomor Surat dan Tanggal harus diisi manual di Excel,
  lalu kedua file di-rename sesuai Nomor Surat

Dengan 300+ file, proses ini memakan waktu **2+ hari kerja** 
dengan risiko human error yang tinggi.

## Solution
**Attempt 1 — Full Automation (Gagal)**
Pendekatan awal: ekstraksi teks PDF via Poppler seperti project BPPU.
Masalah: PDF invoice tidak text-selectable → fallback ke OCR.
OCR berhasil ekstrak teks, tapi variasi format penulisan 
Job Description terlalu tinggi untuk di-match via script → 
terlalu banyak false positive/negative.

**Pivot — Partial Automation**
Saat mengerjakan manual, ditemukan pola kritis yang sebelumnya terlewat:
setiap pasangan PDF-Excel memiliki **kode unik yang konsisten di nama file**.

Insight ini menggeser pendekatan dari *content matching* ke *filename matching*:

1. Script otomatis mencocokkan pasangan PDF-Excel via kode unik di nama file
2. Script mengisi Nomor Surat dan Tanggal di Excel secara otomatis
3. Script me-rename kedua file sesuai Nomor Surat
4. User hanya perlu memverifikasi Job Description dan Subtotal secara visual
   — bagian yang memang butuh human judgment

## Impact
|         Metrik           |               Manual           | Semi-Automated         |
|--------------------------|--------------------------------|------------------------|
| File diproses            |             300+ files         | 300+ files             |
| Waktu per file           |             ~70 detik          | ~20 detik              |
| Bottleneck utama         | Penomoran + rename + verifikasi| Verifikasi konten saja |
| Risiko salah urut/rename |                 Tinggi         | Eliminated             |

## Tech Stack
- Python — core automation logic
- PowerShell — eksekusi di environment kantor (no pip install)

## How It Works
```
Folder Input
  ├── PDF Files  ─┐
  └── Excel Files ─┘
          ↓
  [Filename Matching via CP Code]
          ↓
  Paired: PDF ↔ Excel
          ↓
  Auto-fill: Nomor Surat + Tanggal → Excel
          ↓
  Auto-rename: PDF + Excel → format Nomor Surat
          ↓
  User verifies: Job Description + Subtotal
          ↓
  Flag discrepancies → manual note
```

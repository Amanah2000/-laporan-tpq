import { useState, useEffect } from 'react';
import ExcelJS from 'exceljs';
function App() {
  // 1. Ambil data dari localStorage
  const [dataKeuangan, setDataKeuangan] = useState(() => {
    const saved = localStorage.getItem('dataTPQ');
    const data = saved ? JSON.parse(saved) : [];

    // Ini buat ngasih tanggal ke data lama yang kosong
    return data.map(item => ({
      ...item,
      tanggal: item.tanggal || new Date().toLocaleString('id-ID', {
        day: '2-digit', month: '2-digit', year: 'numeric',
        hour: '2-digit', minute: '2-digit'
      })
    }));
  });
  const [jenis, setJenis] = useState('Pemasukan');
  const [keterangan, setKeterangan] = useState('');
  const [jumlah, setJumlah] = useState('');
  const [tanggalManual, setTanggalManual] = useState(new Date().toISOString().split('T')[0]);

  // 2. Simpan otomatis ke localStorage
  useEffect(() => {
    localStorage.setItem('dataTPQ', JSON.stringify(dataKeuangan));
  }, [dataKeuangan]);

  // 3. Fungsi Tambah Data + Tanggal Otomatis
  const tambahData = () => {
    const dataBaru = {
      id: Date.now(),
      tanggal: new Date().toLocaleString('id-ID', {
        day: '2-digit', month: '2-digit', year: 'numeric',
        hour: '2-digit', minute: '2-digit'
      }),
      jenis,
      keterangan,
      jumlah: parseInt(jumlah)

    };
    setDataKeuangan([...dataKeuangan, dataBaru]);
    setKeterangan('');
    setJumlah('');
  }

  // 4. Fungsi HAPUS DATA - YANG KEMARIN KURANG
  // 4. Fungsi HAPUS DATA - VERSI AMAN VERCEL
  const hapusData = (id) => {
    console.log("ID yg dihapus:", id); // buat ngecek
    setDataKeuangan(dataKeuangan.filter(item => String(item.id) !== String(id)));
  };


  // 5. Hitung total
  const totalPemasukan = dataKeuangan.filter(i => i.jenis === 'Pemasukan').reduce((acc, i) => acc + i.jumlah, 0);
  const totalPengeluaran = dataKeuangan.filter(i => i.jenis === 'Pengeluaran').reduce((acc, i) => acc + i.jumlah, 0);
  const saldo = totalPemasukan - totalPengeluaran;

  // 6. Fungsi DOWNLOAD EXCEL/CSV
  const downloadExcel = async () => { // <-- INI WAJIB ADA "async"
    // ... isi kodenya

    const totalPemasukan = dataKeuangan.filter(item => item.jenis === 'Pemasukan').reduce((sum, item) => sum + item.jumlah, 0);
    const totalPengeluaran = dataKeuangan.filter(item => item.jenis === 'Pengeluaran').reduce((sum, item) => sum + item.jumlah, 0);
    const saldoAkhir = totalPemasukan - totalPengeluaran;
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Keuangan TPQ');
    // 1. Tambah Header
    const headerRow = worksheet.addRow(['Tanggal', 'Jam', 'Jenis', 'Keterangan', 'Jumlah']);

    // Style Header: Biru + Putih Tebal
    headerRow.eachCell((cell) => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2563EB' } };
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
    });

    // 2. Tambah Data
    dataKeuangan.forEach(item => {
      const [tanggal, jam] = item.tanggal.split(', ');
      const row = worksheet.addRow([tanggal, jam, item.jenis, item.keterangan, item.jumlah]);
      row.getCell(5).numFmt = '#,##0'; // Format angka
      row.getCell(5).alignment = { horizontal: 'right' };
    });
    // 3. Tambah Total - VERSI RAPI
    // 3. Tambah Total - POSISI PASTI DI KOLOM E
    worksheet.addRow([]);
    const rowPemasukan = worksheet.addRow(['TOTAL PEMASUKAN', '', '', '', totalPemasukan]);
    const rowPengeluaran = worksheet.addRow(['TOTAL PENGELUARAN', '', '', '', totalPengeluaran]);
    const rowSaldo = worksheet.addRow(['SALDO AKHIR', '', '', '', saldoAkhir]);

    // Style Baris Total: Tebal + Rata Kanan
    [rowPemasukan, rowPengeluaran, rowSaldo].forEach(row => {
      row.font = { bold: true };
      row.getCell(5).alignment = { horizontal: 'right' }; // Kolom E
    });
    // Style Baris Total: Tebal
    // 3. Tambah Total - DIPAKSA RATA KAN KOLOM E
    // Style Baris Total: Tebal
    // Rata kanan semua di Kolom 5 = Kolom E
    worksheet.columns = [
      { width: 12 }, { width: 8 }, { width: 12 }, { width: 25 }, { width: 15 }
    ];

    // 5. Download
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'Laporan_Keuangan_TPQ.xlsx';
    a.click();
  }
  return (
    <div style={{ padding: '20px', fontFamily: 'Arial' }}>
      <h1>Dashboard Keuangan TPQ</h1>

      {/* BOX TOTAL DI ATAS */}
      <div style={{ display: 'flex', gap: '20px', marginBottom: '20px' }}>
        <div style={{ border: '2px solid green', padding: '10px', borderRadius: '8px', width: '147px', }}>
          <p>Total Pemasukan</p>
          <h2 style={{ color: 'green', margin: 0 }}>Rp {totalPemasukan.toLocaleString('id-ID')}</h2>
        </div>
        <div style={{ border: '2px solid red', padding: '10px', borderRadius: '8px', width: '147px' }}>
          <p>Total Pengeluaran</p>
          <h2 style={{ color: 'red', margin: 0 }}>Rp {totalPengeluaran.toLocaleString('id-ID')}</h2>
        </div>
        <div style={{ border: '2px solid blue', padding: '10px', borderRadius: '8px', width: '147px' }}>
          <p>Saldo Akhir</p>
          <h2 style={{ color: 'blue', margin: 0 }}>Rp {saldo.toLocaleString('id-ID')}</h2>
        </div>
      </div>

      {/* FORM INPUT */}

      <div style={{ marginBottom: '20px' }}>
        <select value={jenis} onChange={(e) => setJenis(e.target.value,)}>
          <option>Pemasukan</option>
          <option>Pengeluaran</option>
        </select>
        <button onClick={downloadExcel} style={{ backgroundColor: 'green', color: 'white', marginInlineStart: `345px`, }}>
          Download Excel
        </button>
      </div>

      <div>
        <input
          type="date"
          value={tanggalManual}
          onChange={(e) => setTanggalManual(e.target.value)}
        />
        <input placeholder="Keterangan" value={keterangan} onChange={(e) => setKeterangan(e.target.value)} style={{ margin: '0 5px' }} />
        <input placeholder="Jumlah" type="number" value={jumlah} onChange={(e) => setJumlah(e.target.value)} style={{ margin: '0 5px' }} />
        <button onClick={tambahData}>Tambah</button>


      </div>

      {/* TABEL DATA + TOMBOL HAPUS */}
      <table
        border="1"
        style={{
          width: '42%',
          textAlign: 'center',
        }}
          >
        <thead>

          <tr>
            <th style={{ padding: '8px', border: '1px solid #ddd', backgroundColor: '#2196F3', color: 'white', minWidth: '50px' }}>Tanggal</th>
            <th style={{ padding: '8px', border: '1px solid #ddd', backgroundColor: '#2196F3', color: 'white', minWidth: '50px' }}>Jenis</th>
            <th style={{ padding: '8px', border: '1px solid #ddd', backgroundColor: '#2196F3', color: 'white', minWidth: '50px' }}>Keterangan</th>
            <th style={{ padding: '8px', border: '1px solid #ddd', backgroundColor: '#2196F3', color: 'white', minWidth: '50px' }}>Jumlah</th>
            <th style={{ padding: '8px', border: '1px solid #ddd', backgroundColor: '#2196F3', color: 'white', minWidth: '50px' }}>Aksi</th>
          </tr>
        </thead>

        <tbody>
          {dataKeuangan.map((item, index) => (
            <tr key={item.id}>
              <td style={{ padding: '8px', border: '1px solid #ddd' }}>{item.tanggal}</td>

              <td style={{
                padding: '8px',
                border: '1px solid #ddd',
                color: item.jenis === 'Pemasukan' ? 'green' : 'red',
                fontWeight: 'bold',
                textAlign: 'center'
              }}>
                {item.jenis}
              </td>

              <td style={{ padding: '8px', border: '1px solid #ddd' }}>{item.keterangan}</td>

              <td style={{ padding: '8px', border: '1px solid #ddd', textAlign: 'right' }}>Rp {item.jumlah.toLocaleString()}</td>

              <td style={{ padding: '8px', border: '1px solid #ddd', textAlign: 'center' }}>
                <button onClick={() => hapusData(item.id)} style={{ backgroundColor: 'red', color: 'white', border: 'none', padding: '4px 8px' }}>Hapus</button>
              </td>
            </tr>
          ))}
        </tbody>


      </table>
    </div>
  )
}
/* Tambahkan ini di paling bawah file App.js */

export default App

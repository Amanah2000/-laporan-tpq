import { useState, useEffect } from 'react';
import ExcelJS from 'exceljs';
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, PieChart, Pie, Cell } from 'recharts';
function App() {
  // 1. Ambil data dari localStorage
// 1. Ambil data dari localStorage pas pertama kali buka
const [dataKeuangan, setDataKeuangan] = useState(() => {
  const dataTersimpan = localStorage.getItem('dataKeuanganTPQ');
  return dataTersimpan? JSON.parse(dataTersimpan) : [];
});

// 2. Ini buat ngasih tanggal ke data lama yg kosong
useEffect(() => {
  setDataKeuangan(prevData => 
    prevData.map(item => ({
     ...item,
      tanggal: item.tanggal || new Date().toLocaleString('id-ID', {
        day: '2-digit', month: '2-digit', year: 'numeric',
        hour: '2-digit', minute: '2-digit'
      })
    }))
  );
}, []); // [] artinya cuma jalan 1x pas pertama buka

// 3. Ini buat nyimpen otomatis setiap ada perubahan
useEffect(() => {
  localStorage.setItem('dataKeuanganTPQ', JSON.stringify(dataKeuangan));
}, [dataKeuangan]);
  const [jenis, setJenis] = useState('Pemasukan');
  const [keterangan, setKeterangan] = useState('');
  const [jumlah, setJumlah] = useState('');
  const [tanggalManual, setTanggalManual] = useState(new Date().toISOString().split('T')[0]);

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
  // Data untuk Pie Chart Pengeluaran
  const COLORS = ['#f44336', '#ff9800', '#ffc107', '#8bc34a', '#00bcd4'];
  const dataPie = dataKeuangan
    .filter(item => item.jenis === 'Pengeluaran')
    .reduce((acc, item) => {
      const existing = acc.find(i => i.name === item.keterangan);
      if (existing) {
        existing.value += item.jumlah;
      } else {
        acc.push({ name: item.keterangan, value: item.jumlah });
      }
      return acc;
      
    }, []);
// Data untuk Grafik Batang Bulanan
const dataGrafik = dataKeuangan.reduce((acc, item) => {
  const bulan = new Date(item.tanggal.split(',')[0].split('/').reverse().join('-')).toLocaleString('id-ID', { month: 'short', year: 'numeric' });
  const existing = acc.find(i => i.name === bulan);
  if (existing) {
    if (item.jenis === 'Pemasukan') existing.pemasukan += item.jumlah;
    if (item.jenis === 'Pengeluaran') existing.pengeluaran += item.jumlah;
  } else {
    acc.push({
      name: bulan,
      pemasukan: item.jenis === 'Pemasukan'? item.jumlah : 0,
      pengeluaran: item.jenis === 'Pengeluaran'? item.jumlah : 0,
    });
  }
  return acc;
}, []);


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
      <div className="ringkasan" style={{ display: 'flex', gap: '10px', width: '100%' }}>
        <div style={{ border: '2px solid green', padding: '15px', borderRadius: '8px', flex: 1 }}>
          <div>Total Pemasukan</div>
          <div style={{ color: 'green', fontSize: '24px', fontWeight: 'bold' }}>Rp {totalPemasukan.toLocaleString('id-ID')}</div>
        </div>
        <div style={{ border: '2px solid red', padding: '15px', borderRadius: '8px', flex: 1 }}>
          <div>Total Pengeluaran</div>
          <div style={{ color: 'red', fontSize: '24px', fontWeight: 'bold' }}>Rp {totalPengeluaran.toLocaleString('id-ID')}</div>
        </div>
        <div style={{ border: '2px solid blue', padding: '15px', borderRadius: '8px', flex: 1 }}>
          <div>Saldo Akhir</div>
          <div style={{ color: 'blue', fontSize: '24px', fontWeight: 'bold' }}>Rp {saldo.toLocaleString('id-ID')}</div>
        </div>
      </div>

      {/* FORM INPUT */}

      <div style={{ marginBottom: '20px' }}>
        <select value={jenis} onChange={(e) => setJenis(e.target.value,)}>
          <option>Pemasukan</option>
          <option>Pengeluaran</option>
        </select>
        <input
          type="date"
          value={tanggalManual}
          onChange={(e) => setTanggalManual(e.target.value)}
        />
        <input placeholder="Keterangan" value={keterangan} onChange={(e) => setKeterangan(e.target.value)} style={{ margin: '5px' }} />
        <input placeholder="Jumlah" type="number" value={jumlah} onChange={(e) => setJumlah(e.target.value)} style={{ margin: '5px' }} />
        <button onClick={tambahData}>Tambah</button>
        <button onClick={downloadExcel} style={{ backgroundColor: 'green', color: 'white', marginInlineStart: `5px`, }}>
          Download Excel
        </button>
      </div>


      {/* TABEL DATA + TOMBOL HAPUS */}
      <table
        border="1"
        style={{
          width: '100%',
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
              <td style={{ padding: '8px', border: '1px solid #ddd', textAlign: 'center' }}>Rp {item.jumlah.toLocaleString()}</td>
              <td style={{ padding: '8px', border: '1px solid #ddd', textAlign: 'center' }}>
                <button onClick={() => hapusData(item.id)} style={{ backgroundColor: 'red', color: 'white', border: 'none', padding: '4px 8px' }}>Hapus</button>
              </td>
            </tr>
          ))}
        </tbody>
      </table>
<h2 style={{ marginTop: '30px', textAlign: 'center' }}>Grafik Pengeluaran per Keterangan</h2>
<div style={{ width: '100%', height: 300, marginTop: '20px', backgroundColor: '#fff', padding: '10px', borderRadius: '8px', boxShadow: '0 2px 5px rgba(0,0,0,0.1)' }}>
  <ResponsiveContainer width="100%" height="100%">
    <PieChart>
      <Pie
        data={dataPie}
        cx="50%"
        cy="50%"
        labelLine={false}
        label={({ name, percent }) => `${name}: ${(percent * 100).toFixed(0)}%`}
        outerRadius={80}
        dataKey="value"
      >
        {dataPie.map((entry, index) => (
          <Cell key={`cell-${index}`} fill={COLORS[index % COLORS.length]} />
        ))}
      </Pie>
      <Tooltip formatter={(value) => `Rp ${value.toLocaleString('id-ID')}`} />
      <Legend />
    </PieChart>
  </ResponsiveContainer>
</div>
<h2 style={{ marginTop: '30px', textAlign: 'center' }}>Grafik Pemasukan vs Pengeluaran per Bulan</h2>
<div style={{ width: '100%', height: 350, marginTop: '20px', backgroundColor: '#fff', padding: '10px', borderRadius: '8px', boxShadow: '0 2px 5px rgba(0,0,0,0.1)' }}>
  <ResponsiveContainer width="100%" height="100%">
    <BarChart data={dataGrafik}>
      <CartesianGrid strokeDasharray="3 3" />
      <XAxis dataKey="name" />
      <YAxis tickFormatter={(value) => `Rp ${value/1000}rb`} />
      <Tooltip formatter={(value) => `Rp ${value.toLocaleString('id-ID')}`} />
      <Legend />
      <Bar dataKey="pemasukan" fill="#4CAF50" name="Pemasukan" />
      <Bar dataKey="pengeluaran" fill="#f44336" name="Pengeluaran" />
    </BarChart>
  </ResponsiveContainer>
</div>
    </div>
  )
}
/* Tambahkan ini di paling bawah file App.js */

export default App

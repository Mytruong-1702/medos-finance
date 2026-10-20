import React, { useState } from 'react';
import * as XLSX from 'xlsx';

export default function LocTrungThuChi() {
  const [danhSachGiaoDich, setDanhSachGiaoDich] = useState([]);
  const [tenFile, setTenFile] = useState('');
  const [tieuChiLoc, setTieuChiLoc] = useState('so_hoa_don_noi_dung'); // Mặc định chọn tiêu chí lọc mới
  const [danhSachTrung, setDanhSachTrung] = useState([]);
  const [daTaiLen, setDaTaiLen] = useState(false);

  // Đọc file Excel khi người dùng tải lên
  const handleFileUpload = (e) => {
    const file = e?.target?.files?.[0];
    if (!file) return;
    setTenFile(file.name);

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        let allTransactions = [];

        const sheetNames = workbook?.SheetNames || [];
        sheetNames.forEach((sheetName) => {
          if (!sheetName) return;
          const upperSheetName = String(sheetName).toUpperCase();
          if (upperSheetName.includes('TỔNG')) return;

          const worksheet = workbook.Sheets[sheetName];
          if (!worksheet) return;

          const jsonSheet = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
          if (!Array.isArray(jsonSheet) || jsonSheet.length === 0) return;

          let headerIdx = -1;
          for (let i = 0; i < jsonSheet.length; i++) {
            const row = jsonSheet[i];
            if (!row || !Array.isArray(row)) continue;
            
            const rowStr = row.map(cell => String(cell ?? '')).join(' ').toUpperCase();
            if (rowStr.includes('NỘI DUNG') && (rowStr.includes('SỐ TIỀN') || rowStr.includes('CHI'))) {
              headerIdx = i;
              break;
            }
          }

          if (headerIdx !== -1 && jsonSheet[headerIdx]) {
            const rawHeader = jsonSheet[headerIdx];
            const headers = Array.isArray(rawHeader) ? rawHeader.map(h => String(h ?? '').trim().toUpperCase()) : [];
            
            const idxNgay = headers.findIndex(h => h && h.includes('NGÀY'));
            const idxNoiDung = headers.findIndex(h => h && h.includes('NỘI DUNG'));
            const idxSoTien = headers.findIndex(h => h && (h.includes('SỐ TIỀN') || h.includes('CHI')));
            const idxGhiChu = headers.findIndex(h => h && h.includes('GHI CHÚ'));

            for (let i = headerIdx + 1; i < jsonSheet.length; i++) {
              const row = jsonSheet[i];
              if (!row || !Array.isArray(row)) continue;

              const noiDung = idxNoiDung !== -1 ? row[idxNoiDung] : null;
              const soTien = idxSoTien !== -1 ? row[idxSoTien] : null;

              if (noiDung !== null && noiDung !== undefined && String(noiDung).trim() !== '' && soTien !== null && soTien !== undefined && String(soTien).trim() !== '' && !isNaN(Number(soTien))) {
                let ngayVal = idxNgay !== -1 ? row[idxNgay] : '';
                
                if (typeof ngayVal === 'number') {
                  const dateObj = new Date((ngayVal - (25567 + 2)) * 86400 * 1000);
                  ngayVal = dateObj.toLocaleDateString('vi-VN');
                } else if (ngayVal) {
                  ngayVal = String(ngayVal).trim();
                }

                const ghiChuVal = idxGhiChu !== -1 ? String(row[idxGhiChu] ?? '').trim() : '';

                allTransactions.push({
                  id: `${String(sheetName).trim()}_${i}`,
                  thang: String(sheetName).trim(),
                  ngay: ngayVal,
                  noiDung: String(noiDung).trim(),
                  soTien: Number(soTien),
                  ghiChu: ghiChuVal
                });
              }
            }
          }
        });

        setDanhSachGiaoDich(allTransactions);
        setDaTaiLen(true);
        timKiemTrungLop(allTransactions, tieuChiLoc);
      } catch (err) {
        alert('Lỗi đọc file Excel: ' + (err?.message || err));
      }
    };
    reader.readAsArrayBuffer(file);
  };

  // Trích xuất số hóa đơn từ chuỗi ghi chú (tìm dãy số có từ 4 đến 12 chữ số)
  const trichXuatSoHoaDon = (ghiChu) => {
    if (!ghiChu) return '';
    const match = ghiChu.match(/\b\d{4,12}\b/);
    return match ? match[0] : '';
  };

  // Thuật toán tìm kiếm trùng lặp tích hợp tiêu chí lọc theo Số hóa đơn & Nội dung
  const timKiemTrungLop = (data, criteria) => {
    const groups = {};

    (data || []).forEach((item) => {
      let key = '';
      const cleanContent = String(item?.noiDung || '').toLowerCase().replace(/\s+/g, ' ').trim();
      const soHoaDon = trichXuatSoHoaDon(item?.ghiChu);
      
      if (criteria === 'so_hoa_don_noi_dung') {
        // Tiêu chí mới: Trùng theo Số hóa đơn VÀ Nội dung (chỉ gom nhóm khi có chung Số hóa đơn và Nội dung)
        // Nếu không tìm thấy số hóa đơn trong ghi chú, dùng nội dung và số tiền làm khóa thay thế để tránh gom nhầm rác trống
        key = soHoaDon ? `SHD_${soHoaDon}_ND_${cleanContent}` : `ND_${cleanContent}_TIEN_${item?.soTien || 0}`;
      } else if (criteria === 'ngay_noi_dung_tien') {
        key = `${item?.ngay || ''}_${cleanContent}_${item?.soTien || 0}`;
      } else {
        key = `${cleanContent}_${item?.soTien || 0}`;
      }

      if (!groups[key]) {
        groups[key] = [];
      }
      groups[key].push(item);
    });

    const duplicateGroups = Object.keys(groups)
      .filter(key => groups[key].length > 1)
      .map(key => ({
        key,
        items: groups[key]
      }));

    setDanhSachTrung(duplicateGroups);
  };

  const handleDoiTieuChi = (e) => {
    const val = e.target.value;
    setTieuChiLoc(val);
    if (danhSachGiaoDich.length > 0) {
      timKiemTrungLop(danhSachGiaoDich, val);
    }
  };

  const handleXoaGiaoDichKhoiDanhSachChiTiet = (id) => {
    const updated = danhSachGiaoDich.filter(item => item.id !== id);
    setDanhSachGiaoDich(updated);
    timKiemTrungLop(updated, tieuChiLoc);
  };

  return (
    <div style={{ maxWidth: '1400px', margin: '0 auto', fontFamily: '"Segoe UI", Roboto, sans-serif', padding: '30px 20px', backgroundColor: '#f2f7f4', minHeight: '100vh', boxSizing: 'border-box' }}>
      
      {/* HEADER TITLE */}
      <div style={{ backgroundColor: '#ffffff', padding: '24px 30px', borderRadius: '14px', border: '1px solid #d4e2d8', boxShadow: '0 4px 15px rgba(0,0,0,0.02)', marginBottom: '24px' }}>
        <h2 style={{ margin: '0 0 8px 0', color: '#3b5343', fontSize: '22px', fontWeight: '800' }}>
          🔍 Tính năng Lọc Trùng Thu Chi
        </h2>
        <p style={{ margin: 0, color: '#688271', fontSize: '13.5px' }}>
          Tải lên file Excel thống kê thu chi để tự động quét trùng theo Số hóa đơn trong ghi chú và Nội dung thanh toán.
        </p>
      </div>

      {/* KHUNG TẢI FILE VÀ CẤU HÌNH */}
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '20px', marginBottom: '24px' }}>
        
        {/* Tải file */}
        <div style={{ backgroundColor: '#ffffff', padding: '20px 24px', borderRadius: '12px', border: '1px solid #d4e2d8', display: 'flex', flexDirection: 'column', justifyContent: 'center' }}>
          <label style={{ fontSize: '13px', fontWeight: '700', color: '#3b5343', marginBottom: '8px', display: 'block' }}>
            📁 Tải lên tệp Excel thu chi:
          </label>
          <input 
            type="file" 
            accept=".xlsx, .xls" 
            onChange={handleFileUpload} 
            style={{ fontSize: '13px', color: '#4f6d5a', padding: '8px', border: '1px dashed #b3c4b9', borderRadius: '8px', backgroundColor: '#f7faf8', cursor: 'pointer' }}
          />
          {tenFile && <span style={{ fontSize: '12px', color: '#4f6d5a', marginTop: '6px', fontWeight: '600' }}>Đã tải: {tenFile} ({danhSachGiaoDich.length} giao dịch)</span>}
        </div>

        {/* Tiêu chí lọc */}
        <div style={{ backgroundColor: '#ffffff', padding: '20px 24px', borderRadius: '12px', border: '1px solid #d4e2d8', display: 'flex', flexDirection: 'column', justifyContent: 'center' }}>
          <label style={{ fontSize: '13px', fontWeight: '700', color: '#3b5343', marginBottom: '8px', display: 'block' }}>
            ⚙️ Tiêu chí phát hiện trùng lặp:
          </label>
          <select 
            value={tieuChiLoc} 
            onChange={handleDoiTieuChi}
            style={{ padding: '10px 12px', borderRadius: '8px', border: '1px solid #b3c4b9', fontSize: '13.5px', color: '#2c3e32', backgroundColor: '#f7faf8', outline: 'none', fontWeight: '600' }}
          >
            <option value="so_hoa_don_noi_dung">Trùng theo Số hóa đơn (trong Ghi chú) & Nội dung</option>
            <option value="ngay_noi_dung_tien">Trùng tuyệt đối (Cùng Ngày + Cùng Nội dung + Cùng Số tiền)</option>
            <option value="noi_dung_tien">Trùng theo Nội dung & Số tiền (Bất kể ngày/tháng khác nhau)</option>
          </select>
        </div>

      </div>

      {/* KẾT QUẢ QUÉT TRÙNG */}
      {daTaiLen && (
        <div style={{ backgroundColor: '#ffffff', padding: '24px 30px', borderRadius: '14px', border: '1px solid #d4e2d8', boxShadow: '0 4px 15px rgba(0,0,0,0.02)' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px', borderBottom: '1px solid #e2ece5', paddingBottom: '12px' }}>
            <h3 style={{ margin: 0, fontSize: '16px', color: '#3b5343', fontWeight: '800' }}>
              📊 Kết quả phát hiện: <span style={{ color: danhSachTrung.length > 0 ? '#b5651d' : '#32523b' }}>{danhSachTrung.length} nhóm nghi ngờ trùng lặp</span>
            </h3>
            <span style={{ fontSize: '12.5px', backgroundColor: '#eaf2ec', color: '#3b5343', padding: '4px 10px', borderRadius: '6px', fontWeight: '700' }}>
              Tổng số giao dịch đã quét: {danhSachGiaoDich.length}
            </span>
          </div>

          {danhSachTrung.length === 0 ? (
            <div style={{ textAlign: 'center', padding: '40px', color: '#688271', fontSize: '14px', fontWeight: '600' }}>
              🎉 Tuyệt vời! Không phát hiện khoản chi thu nào bị trùng lặp theo tiêu chí này.
            </div>
          ) : (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '16px' }}>
              {danhSachTrung.map((group, idx) => {
                const soHDGroup = trichXuatSoHoaDon(group.items[0].ghiChu);
                return (
                  <div key={idx} style={{ border: '1px solid #f2d9b1', borderRadius: '10px', backgroundColor: '#fffdf9', overflow: 'hidden' }}>
                    <div style={{ backgroundColor: '#fff8eb', padding: '10px 16px', borderBottom: '1px solid #f2d9b1', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                      <span style={{ fontSize: '12.5px', fontWeight: '700', color: '#966712' }}>
                        Nhóm trùng #{idx + 1} — {soHDGroup ? `Số hóa đơn: ${soHDGroup}` : 'Trùng nội dung'} ({group.items.length} bản ghi)
                      </span>
                      <span style={{ fontSize: '11.5px', color: '#8c703b' }}>Nội dung: "{group.items[0].noiDung}"</span>
                    </div>

                    <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12.5px' }}>
                      <thead>
                        <tr style={{ backgroundColor: '#f9f6ef', color: '#7a5d2b', borderBottom: '1px solid #f2d9b1', textAlign: 'left' }}>
                          <th style={{ padding: '8px 12px' }}>Tháng</th>
                          <th style={{ padding: '8px 12px' }}>Ngày</th>
                          <th style={{ padding: '8px 12px' }}>Nội dung chi tiết</th>
                          <th style={{ padding: '8px 12px' }}>Số tiền (VNĐ)</th>
                          <th style={{ padding: '8px 12px' }}>Ghi chú (Số hóa đơn)</th>
                          <th style={{ padding: '8px 12px', textAlign: 'center' }}>Thao tác xử lý</th>
                        </tr>
                      </thead>
                      <tbody>
                        {group.items.map((item, subIdx) => (
                          <tr key={subIdx} style={{ borderBottom: '1px solid #f7f2e8' }}>
                            <td style={{ padding: '8px 12px', fontWeight: '600', color: '#3b5343' }}>{item.thang}</td>
                            <td style={{ padding: '8px 12px' }}>{item.ngay || '-'}</td>
                            <td style={{ padding: '8px 12px', fontWeight: '600', color: '#2c3e32' }}>{item.noiDung}</td>
                            <td style={{ padding: '8px 12px', fontWeight: '800', color: '#966712' }}>{item.soTien.toLocaleString('vi-VN')}</td>
                            <td style={{ padding: '8px 12px', fontStyle: 'italic', color: '#688271' }}>{item.ghiChu || '-'}</td>
                            <td style={{ padding: '8px 12px', textAlign: 'center' }}>
                              <button 
                                onClick={() => handleXoaGiaoDichKhoiDanhSachChiTiet(item.id)}
                                style={{ backgroundColor: '#f5dede', color: '#a84343', border: 'none', padding: '4px 10px', borderRadius: '5px', fontSize: '11px', fontWeight: 'bold', cursor: 'pointer' }}
                                title="Loại bỏ bản ghi trùng thừa này"
                              >
                                Xóa bản ghi này
                              </button>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                );
              })}
            </div>
          )}
        </div>
      )}

    </div>
  );
}
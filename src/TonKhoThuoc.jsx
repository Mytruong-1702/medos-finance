import React, { useState } from 'react';
import * as XLSX from 'xlsx';

export default function TonKhoThuoc() {
  const [dataTonKho, setDataTonKho] = useState([]);
  const [tieuDeDonVi, setTieuDeDonVi] = useState('SỞ Y TẾ CÀ MAU - PKĐK Y ĐỨC SÀI GÒN');
  const [ngayThongKe, setNgayThongKe] = useState('NGÀY 30/09/2026');
  const [tuKhoaTimKiem, setTuKhoaTimKiem] = useState('');
  const [boLocTrangThai, setBoLocTrangThai] = useState('ALL');
  const [loading, setLoading] = useState(false);

  // Hàm chuyển đổi định dạng số kiểu Việt Nam (VD: 1.950 hoặc 7.968.000,0000) thành kiểu Number chuẩn
  const parseVietnameseNumber = (val) => {
    if (val === null || val === undefined) return 0;
    if (typeof val === 'number') return val;
    let s = String(val).trim();
    if (!s) return 0;
    
    // Nếu có chứa dấu phẩy (phần thập phân)
    if (s.includes(',')) {
      s = s.replace(/\./g, '').replace(',', '.');
    } else {
      // Nếu không có dấu phẩy nhưng có nhiều dấu chấm (vd: 7.968.000) hoặc dấu chấm phân cách hàng nghìn (1.950)
      s = s.replace(/\./g, '');
    }
    const num = parseFloat(s);
    return isNaN(num) ? 0 : num;
  };

  // Đọc file Excel Thống kê tồn kho hệ thống
  const handleUploadTonKhoFile = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setLoading(true);
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        let headerIndex = -1;
        let ngayBaoCao = 'NGÀY 30/09/2026';
        let donViBaoCao = 'SỞ Y TẾ CÀ MAU - PKĐK Y ĐỨC SÀI GÒN';

        for (let i = 0; i < Math.min(rows.length, 15); i++) {
          const rowStr = rows[i].map(c => String(c || '')).join(' ');
          if (rowStr.toUpperCase().includes('SỞ Y TẾ') || rowStr.toUpperCase().includes('PKĐK')) {
            donViBaoCao = rowStr;
          }
          if (rowStr.includes('NGÀY')) {
            ngayBaoCao = rowStr;
          }
          if (rows[i].some(val => val && (String(val).includes('Mã VT') || String(val).includes('Tên vật tư')))) {
            headerIndex = i;
            break;
          }
        }

        if (headerIndex === -1) {
          alert('Không tìm thấy dòng tiêu đề bảng hợp lệ trong file Excel!');
          setLoading(false);
          return;
        }

        setTieuDeDonVi(donViBaoCao);
        setNgayThongKe(ngayBaoCao);

        const headers = rows[headerIndex];
        const dataRows = rows.slice(headerIndex + 1);

        const formattedData = dataRows.map((row, idx) => {
          const obj = {};
          headers.forEach((h, colIdx) => {
            if (h) obj[String(h).trim()] = row[colIdx];
          });
          const soLuongHeThong = parseVietnameseNumber(obj['Số lượng']);
          const donGiaVal = parseVietnameseNumber(obj['Đơn giá']);

          return {
            stt: row[0] || (idx + 1),
            maVt: String(obj['Mã VT'] || obj['MA_VT'] || '').trim(),
            tenVatTu: String(obj['Tên vật tư'] || obj['TEN_VAT_TU'] || '').trim(),
            hoatChat: String(obj['Hoạt chất'] || '').trim(),
            hamLuong: String(obj['Hàm lượng'] || '').trim(),
            dvt: String(obj['ĐVT'] || '').trim(),
            soLo: String(obj['Số lô'] || '').trim(),
            soThau: String(obj['Số thầu'] || '').trim(),
            ngayHetHan: String(obj['Ngày hết hạn'] || '').trim(),
            donGia: donGiaVal,
            soLuongHeThong: soLuongHeThong,
            soLuongThucKiem: soLuongHeThong,
          };
        }).filter(item => item.tenVatTu && item.tenVatTu !== 'undefined');

        setDataTonKho(formattedData);
        alert(`Đã phân tích thành công ${formattedData.length} mặt hàng tồn kho!`);
      } catch (err) {
        console.error(err);
        alert('Lỗi đọc và phân tích file Excel tồn kho!');
      } finally {
        setLoading(false);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  // Cập nhật số lượng thực kiểm
  const handleThucKiemChange = (index, value) => {
    const newTonKho = [...dataTonKho];
    newTonKho[index].soLuongThucKiem = value === '' ? '' : Number(value);
    setDataTonKho(newTonKho);
  };

  // Tính toán chênh lệch và thành tiền
  const processedData = dataTonKho.map(item => {
    const thucKiem = item.soLuongThucKiem === '' ? item.soLuongHeThong : Number(item.soLuongThucKiem);
    const chenhLech = thucKiem - item.soLuongHeThong;

    let lechThieu = '';
    let lechDu = '';
    let thanhTienThieu = '';
    let thanhTienDu = '';

    if (chenhLech < 0) {
      lechThieu = chenhLech;
      thanhTienThieu = chenhLech * item.donGia;
    } else if (chenhLech > 0) {
      lechDu = chenhLech;
      thanhTienDu = chenhLech * item.donGia;
    }

    return {
      ...item,
      thucKiemNumber: thucKiem,
      lechThieu,
      lechDu,
      thanhTienThieu,
      thanhTienDu,
      thanhTienHeThong: item.soLuongHeThong * item.donGia
    };
  });

  // Lọc dữ liệu
  const filteredData = processedData.filter(item => {
    const keyword = tuKhoaTimKiem.toLowerCase();
    const matchKeyword = (
      item.tenVatTu.toLowerCase().includes(keyword) ||
      item.maVt.toLowerCase().includes(keyword) ||
      item.hoatChat.toLowerCase().includes(keyword) ||
      item.soLo.toLowerCase().includes(keyword)
    );

    if (boLocTrangThai === 'ALL') return matchKeyword;
    if (boLocTrangThai === 'THIEU') return matchKeyword && item.lechThieu !== '';
    if (boLocTrangThai === 'THUA') return matchKeyword && item.lechDu !== '';
    if (boLocTrangThai === 'KHOP') return matchKeyword && item.lechThieu === '' && item.lechDu === '';
    return matchKeyword;
  });

  // Tổng hợp thống kê
  const tongSoLuongHeThong = processedData.reduce((sum, i) => sum + i.soLuongHeThong, 0);
  const tongThanhTienHeThong = processedData.reduce((sum, i) => sum + i.thanhTienHeThong, 0);
  const tongLechThieu = processedData.reduce((sum, i) => sum + (i.lechThieu !== '' ? i.lechThieu : 0), 0);
  const tongLechDu = processedData.reduce((sum, i) => sum + (i.lechDu !== '' ? i.lechDu : 0), 0);
  const tongThanhTienThieu = processedData.reduce((sum, i) => sum + (i.thanhTienThieu !== '' ? i.thanhTienThieu : 0), 0);
  const tongThanhTienDu = processedData.reduce((sum, i) => sum + (i.thanhTienDu !== '' ? i.thanhTienDu : 0), 0);

  // Xuất Excel
  const exportToExcel = () => {
    if (dataTonKho.length === 0) {
      alert('Không có dữ liệu để xuất!');
      return;
    }
    const exportRows = processedData.map((item, idx) => ({
      'ST': idx + 1,
      'Mã VT': item.maVt,
      'Tên vật tư': item.tenVatTu,
      'Hoạt chất': item.hoatChat,
      'Hàm lượng': item.hamLuong,
      'ĐVT': item.dvt,
      'Số lô': item.soLo,
      'Số thầu': item.soThau,
      'Ngày hết hạn': item.ngayHetHan,
      'Đơn giá': item.donGia,
      'Số lượng': item.soLuongHeThong,
      'Thực kiểm': item.thucKiemNumber,
      'Thành tiền': item.thanhTienHeThong,
      'Lệch thiếu': item.lechThieu,
      'Lệch dư': item.lechDu,
      'Thành tiền thiếu': item.thanhTienThieu,
      'Thành tiền dư': item.thanhTienDu
    }));

    const worksheet = XLSX.utils.json_to_sheet(exportRows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "ThongKeTonKiemKe");
    XLSX.writeFile(workbook, "ThongKeTon_KiemKe_medOS.xlsx");
  };

  return (
    <div style={{ width: '100%', maxWidth: '100%', padding: '15px', boxSizing: 'border-box', backgroundColor: '#ffffff', borderRadius: '8px', boxShadow: '0 4px 12px rgba(0,0,0,0.05)' }}>
      <div style={{ textAlign: 'center', marginBottom: '15px' }}>
        <h4 style={{ margin: 0, color: '#1e293b', fontSize: '14px' }}>{tieuDeDonVi}</h4>
        <h2 style={{ color: '#0284c7', margin: '4px 0', fontSize: '18px' }}>THỐNG KÊ TỒN & KIỂM KÊ THỰC TẾ</h2>
        <div style={{ fontWeight: 'bold', color: '#16a34a', fontSize: '13px' }}>📅 {ngayThongKe}</div>
      </div>

      <div style={{ display: 'flex', gap: '12px', alignItems: 'center', marginBottom: '15px', flexWrap: 'wrap', backgroundColor: '#f0f9ff', padding: '12px', borderRadius: '6px' }}>
        <div>
          <label style={{ fontWeight: 'bold', display: 'block', marginBottom: '4px', fontSize: '12px' }}>Tải lên file Thống Kê Tồn (Hệ thống):</label>
          <input type="file" accept=".xlsx, .xls" onChange={handleUploadTonKhoFile} style={{ fontSize: '11px' }} />
        </div>
        <div>
          <label style={{ fontWeight: 'bold', display: 'block', marginBottom: '4px', fontSize: '12px' }}>Tìm kiếm tên thuốc / số lô:</label>
          <input 
            type="text" 
            placeholder="Nhập tên, mã, số lô..." 
            value={tuKhoaTimKiem} 
            onChange={(e) => setTuKhoaTimKiem(e.target.value)}
            style={{ padding: '6px 10px', borderRadius: '4px', border: '1px solid #cbd5e1', width: '200px', fontSize: '12px' }}
          />
        </div>
        <div>
          <label style={{ fontWeight: 'bold', display: 'block', marginBottom: '4px', fontSize: '12px' }}>Bộ lọc:</label>
          <select 
            value={boLocTrangThai} 
            onChange={(e) => setBoLocTrangThai(e.target.value)}
            style={{ padding: '6px 10px', borderRadius: '4px', border: '1px solid #cbd5e1', fontSize: '12px', backgroundColor: '#fff' }}
          >
            <option value="ALL">Tất cả mặt hàng</option>
            <option value="THIEU">🔴 Chỉ hiện Hàng Thiếu</option>
            <option value="THUA">🟡 Chỉ hiện Hàng Dư (Thừa)</option>
            <option value="KHOP">🟢 Chỉ hiện Hàng Khớp</option>
          </select>
        </div>
        <div style={{ marginLeft: 'auto' }}>
          <button onClick={exportToExcel} style={{ padding: '7px 14px', backgroundColor: '#10b981', color: 'white', border: 'none', borderRadius: '6px', fontWeight: 'bold', cursor: 'pointer', fontSize: '12px' }}>
            📥 Xuất Excel Báo Cáo Kiểm Kê
          </button>
        </div>
      </div>

      {loading && <p style={{ textAlign: 'center', fontWeight: 'bold', color: '#0284c7' }}>⏳ Đang phân tích dữ liệu kho...</p>}

      {/* Bảng dữ liệu tối ưu hiển thị, thu gọn cột hoạt chất/hàm lượng */}
      <div style={{ width: '100%', overflowX: 'auto', maxHeight: '650px', border: '1px solid #cbd5e1' }}>
        <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '11.5px', whiteSpace: 'nowrap' }}>
          <thead>
            <tr style={{ backgroundColor: '#1e293b', color: 'white', textAlign: 'center' }}>
              <th style={{ padding: '7px 8px', border: '1px solid #cbd5e1' }}>ST</th>
              <th style={{ padding: '7px 8px', border: '1px solid #cbd5e1' }}>Mã VT</th>
              <th style={{ padding: '7px 8px', border: '1px solid #cbd5e1', textAlign: 'left', minWidth: '160px' }}>Tên vật tư</th>
              <th style={{ padding: '7px 8px', border: '1px solid #cbd5e1', textAlign: 'left', maxWidth: '180px', overflow: 'hidden', textOverflow: 'ellipsis' }}>Hoạt chất</th>
              <th style={{ padding: '7px 8px', border: '1px solid #cbd5e1', maxWidth: '100px', overflow: 'hidden', textOverflow: 'ellipsis' }}>Hàm lượng</th>
              <th style={{ padding: '7px 8px', border: '1px solid #cbd5e1' }}>ĐVT</th>
              <th style={{ padding: '7px 8px', border: '1px solid #cbd5e1' }}>Số lô</th>
              <th style={{ padding: '7px 8px', border: '1px solid #cbd5e1' }}>Số thầu</th>
              <th style={{ padding: '7px 8px', border: '1px solid #cbd5e1' }}>Ngày hết hạn</th>
              <th style={{ padding: '7px 8px', border: '1px solid #cbd5e1', textAlign: 'right' }}>Đơn giá</th>
              <th style={{ padding: '7px 8px', border: '1px solid #cbd5e1', textAlign: 'right' }}>Số lượng</th>
              <th style={{ padding: '7px 8px', border: '1px solid #cbd5e1', textAlign: 'right', backgroundColor: '#0284c7', minWidth: '90px' }}>Thực kiểm</th>
              <th style={{ padding: '7px 8px', border: '1px solid #cbd5e1', textAlign: 'right' }}>Thành tiền</th>
              <th style={{ padding: '7px 8px', border: '1px solid #cbd5e1', textAlign: 'right', color: '#f87171' }}>Lệch thiếu</th>
              <th style={{ padding: '7px 8px', border: '1px solid #cbd5e1', textAlign: 'right', color: '#facc15' }}>Lệch dư</th>
              <th style={{ padding: '7px 8px', border: '1px solid #cbd5e1', textAlign: 'right', color: '#f87171' }}>Thành tiền thiếu</th>
              <th style={{ padding: '7px 8px', border: '1px solid #cbd5e1', textAlign: 'right', color: '#4ade80' }}>Thành tiền dư</th>
            </tr>
          </thead>
          <tbody>
            {filteredData.length > 0 ? (
              filteredData.map((row, idx) => {
                const originalIndex = dataTonKho.findIndex(item => item.maVt === row.maVt && item.soLo === row.soLo);
                const isThieu = row.lechThieu !== '';
                const isDu = row.lechDu !== '';

                return (
                  <tr key={idx} style={{ 
                    borderBottom: '1px solid #cbd5e1', 
                    backgroundColor: isThieu ? '#fff5f5' : isDu ? '#fefce8' : (idx % 2 === 0 ? '#fff' : '#f8fafc') 
                  }}>
                    <td style={{ padding: '6px 8px', border: '1px solid #cbd5e1', textAlign: 'center' }}>{row.stt}</td>
                    <td style={{ padding: '6px 8px', border: '1px solid #cbd5e1', textAlign: 'center' }}>{row.maVt}</td>
                    <td style={{ padding: '6px 8px', border: '1px solid #cbd5e1', fontWeight: '600', color: '#1e293b' }}>{row.tenVatTu}</td>
                    <td style={{ padding: '6px 8px', border: '1px solid #cbd5e1', color: '#4b5563', maxWidth: '180px', overflow: 'hidden', textOverflow: 'ellipsis' }} title={row.hoatChat}>{row.hoatChat}</td>
                    <td style={{ padding: '6px 8px', border: '1px solid #cbd5e1', textAlign: 'center', maxWidth: '100px', overflow: 'hidden', textOverflow: 'ellipsis' }} title={row.hamLuong}>{row.hamLuong}</td>
                    <td style={{ padding: '6px 8px', border: '1px solid #cbd5e1', textAlign: 'center' }}>{row.dvt}</td>
                    <td style={{ padding: '6px 8px', border: '1px solid #cbd5e1', textAlign: 'center', fontWeight: '600' }}>{row.soLo}</td>
                    <td style={{ padding: '6px 8px', border: '1px solid #cbd5e1', textAlign: 'center' }}>{row.soThau}</td>
                    <td style={{ padding: '6px 8px', border: '1px solid #cbd5e1', textAlign: 'center', color: '#dc2626', fontWeight: '600' }}>{row.ngayHetHan}</td>
                    <td style={{ padding: '6px 8px', border: '1px solid #cbd5e1', textAlign: 'right' }}>{row.donGia.toLocaleString()}</td>
                    <td style={{ padding: '6px 8px', border: '1px solid #cbd5e1', textAlign: 'right', fontWeight: 'bold', color: '#0284c7' }}>{row.soLuongHeThong.toLocaleString()}</td>
                    <td style={{ padding: '6px 8px', border: '1px solid #cbd5e1', textAlign: 'center', backgroundColor: '#f0f9ff' }}>
                      <input 
                        type="number" 
                        value={row.soLuongThucKiem} 
                        onChange={(e) => handleThucKiemChange(originalIndex, e.target.value)}
                        style={{ width: '75px', padding: '4px 6px', textAlign: 'center', fontWeight: 'bold', border: '2px solid #0284c7', borderRadius: '4px', backgroundColor: '#fff', fontSize: '12px' }}
                      />
                    </td>
                    <td style={{ padding: '6px 8px', border: '1px solid #cbd5e1', textAlign: 'right', fontWeight: 'bold' }}>{row.thanhTienHeThong.toLocaleString()}</td>
                    <td style={{ padding: '6px 8px', border: '1px solid #cbd5e1', textAlign: 'right', fontWeight: 'bold', color: '#dc2626' }}>{row.lechThieu !== '' ? row.lechThieu.toLocaleString() : ''}</td>
                    <td style={{ padding: '6px 8px', border: '1px solid #cbd5e1', textAlign: 'right', fontWeight: 'bold', color: '#ca8a04' }}>{row.lechDu !== '' ? `+${row.lechDu.toLocaleString()}` : ''}</td>
                    <td style={{ padding: '6px 8px', border: '1px solid #cbd5e1', textAlign: 'right', fontWeight: 'bold', color: '#dc2626' }}>{row.thanhTienThieu !== '' ? row.thanhTienThieu.toLocaleString() : ''}</td>
                    <td style={{ padding: '6px 8px', border: '1px solid #cbd5e1', textAlign: 'right', fontWeight: 'bold', color: '#16a34a' }}>{row.thanhTienDu !== '' ? row.thanhTienDu.toLocaleString() : ''}</td>
                  </tr>
                );
              })
            ) : (
              <tr>
                <td colSpan="17" style={{ textAlign: 'center', padding: '40px', color: '#64748b' }}>
                  {dataTonKho.length === 0 ? "📁 Vui lòng tải lên file Excel Thống kê tồn hệ thống để bắt đầu kiểm kê." : "Không tìm thấy mặt hàng phù hợp."}
                </td>
              </tr>
            )}
          </tbody>
          {dataTonKho.length > 0 && (
            <tfoot>
              <tr style={{ backgroundColor: '#f1f5f9', fontWeight: 'bold', textAlign: 'right' }}>
                <td colSpan="10" style={{ padding: '8px', border: '1px solid #cbd5e1', textAlign: 'center' }}>TỔNG CỘNG</td>
                <td style={{ padding: '8px', border: '1px solid #cbd5e1', color: '#0284c7' }}>{tongSoLuongHeThong.toLocaleString()}</td>
                <td style={{ padding: '8px', border: '1px solid #cbd5e1' }}>-</td>
                <td style={{ padding: '8px', border: '1px solid #cbd5e1' }}>{tongThanhTienHeThong.toLocaleString()}</td>
                <td style={{ padding: '8px', border: '1px solid #cbd5e1', color: '#dc2626' }}>{tongLechThieu !== 0 ? tongLechThieu.toLocaleString() : ''}</td>
                <td style={{ padding: '8px', border: '1px solid #cbd5e1', color: '#ca8a04' }}>{tongLechDu !== 0 ? `+${tongLechDu.toLocaleString()}` : ''}</td>
                <td style={{ padding: '8px', border: '1px solid #cbd5e1', color: '#dc2626' }}>{tongThanhTienThieu !== 0 ? tongThanhTienThieu.toLocaleString() : ''}</td>
                <td style={{ padding: '8px', border: '1px solid #cbd5e1', color: '#16a34a' }}>{tongThanhTienDu !== 0 ? tongThanhTienDu.toLocaleString() : ''}</td>
              </tr>
            </tfoot>
          )}
        </table>
      </div>
    </div>
  );
}
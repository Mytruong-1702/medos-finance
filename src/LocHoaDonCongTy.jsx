import React, { useState, useMemo } from 'react';
import ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';

function LocHoaDonCongTy({ duLieuHoaDon: propsDuLieu }) {
  const [cachedData, setCachedData] = useState(() => {
    try {
      const saved = localStorage.getItem('APP_HOA_DON_CACHE_V2');
      return saved ? JSON.parse(saved) : [];
    } catch {
      return [];
    }
  });

  const [selectedCompany, setSelectedCompany] = useState('ALL');
  const [loadingExport, setLoadingExport] = useState(false);
  const [tenFileTaiLen, setTenFileTaiLen] = useState('Dữ liệu từ hệ thống chính');

  const duLieuHoaDon = useMemo(() => {
    if (propsDuLieu && Array.isArray(propsDuLieu) && propsDuLieu.length > 0) {
      localStorage.setItem('APP_HOA_DON_CACHE_V2', JSON.stringify(propsDuLieu));
      return propsDuLieu;
    }
    return cachedData;
  }, [propsDuLieu, cachedData]);

  const handleDirectFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setTenFileTaiLen(file.name);

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const rowsRaw = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { header: 1, defval: '' });

        let headerRowIndex = 0;
        for (let i = 0; i < rowsRaw.length; i++) {
          if (rowsRaw[i] && rowsRaw[i].some(cell => String(cell).toLowerCase().includes('số hóa đơn') || String(cell).toLowerCase().includes('tên người bán'))) {
            headerRowIndex = i;
            break;
          }
        }

        const headers = rowsRaw[headerRowIndex] || [];
        const dataRows = rowsRaw.slice(headerRowIndex + 1);

        const listHoaDonMoi = dataRows.map((row, idx) => {
          const convertRowToObject = { "STT": idx + 1 };
          headers.forEach((header, index) => { 
            if (header) convertRowToObject[String(header).trim()] = row[index] !== undefined ? row[index] : ''; 
          });

          let soHD = "", tenCty = "", ngayHD = "";
          Object.keys(convertRowToObject).forEach(k => {
            const lk = k.toLowerCase();
            if (lk.includes('số hóa đơn') || lk.includes('so_hoa_don')) soHD = convertRowToObject[k];
            if (lk.includes('tên người bán') || lk.includes('ten_nguoi_ban') || lk.includes('đơn vị bán')) tenCty = convertRowToObject[k];
            if (lk.includes('ngày lập') || lk.includes('ngày hóa đơn') || lk.includes('ngay_lap')) ngayHD = convertRowToObject[k];
          });

          return {
            maHoaDon: String(soHD || '').trim(),
            tenCongTy: String(tenCty || 'Chưa rõ nhà cung cấp').trim(),
            ngayHoaDon: String(ngayHD || '').trim(),
            rowGoc: convertRowToObject
          };
        }).filter(item => item.rowGoc && Object.values(item.rowGoc).some(v => v !== ''));

        if (listHoaDonMoi.length > 0) {
          setCachedData(listHoaDonMoi);
          localStorage.setItem('APP_HOA_DON_CACHE_V2', JSON.stringify(listHoaDonMoi));
          setSelectedCompany('ALL');
        } else {
          alert("⚠ Không tìm thấy dữ liệu hóa đơn trong file này!");
        }
      } catch (err) {
        console.error(err);
        alert("❌ Lỗi đọc file Excel trực tiếp!");
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const danhSachCongTy = useMemo(() => {
    if (!duLieuHoaDon || duLieuHoaDon.length === 0) return ['ALL'];
    const ctySet = new Set();
    duLieuHoaDon.forEach(item => {
      if (item && item.tenCongTy) {
        ctySet.add(item.tenCongTy.trim());
      }
    });
    return ['ALL', ...Array.from(ctySet)];
  }, [duLieuHoaDon]);

  const duLieuDaLoc = useMemo(() => {
    if (!duLieuHoaDon) return [];
    if (selectedCompany === 'ALL') return duLieuHoaDon;
    return duLieuHoaDon.filter(item => item && item.tenCongTy && item.tenCongTy.trim() === selectedCompany);
  }, [duLieuHoaDon, selectedCompany]);

  const tatCaTieuDeCotGoc = useMemo(() => {
    if (!duLieuHoaDon || duLieuHoaDon.length === 0) return [];
    return Object.keys(duLieuHoaDon[0].rowGoc || {});
  }, [duLieuHoaDon]);

  // 🌟 TỰ ĐỘNG PHÁT HIỆN 3 CỘT TIỀN TRONG BẢNG ĐỂ HIỂN THỊ CHÍNH XÁC KHỚP VỚI GIAO DIỆN
  const thongKeTaiChinh = useMemo(() => {
    let col1Val = 0; 
    let col2Val = 0; 
    let col3Val = 0;

    // Lấy tên 3 cột tiền cuối cùng hoặc 3 cột tiền đặc trưng trong bảng
    const columns = tatCaTieuDeCotGoc;
    const moneyColumns = columns.filter(col => {
      const lk = col.toLowerCase();
      return lk.includes('tiền') || lk.includes('thuế') || lk.includes('doanh số') || lk.includes('thanh toán') || lk.includes('chiết khấu');
    });

    // Chọn ra 3 cột tiền quan trọng nhất (ví dụ: Tổng tiền hàng, Tiền chiết khấu, Doanh số bán chưa thuế / Tiền thuế / Thanh toán)
    const targetCol1 = moneyColumns[0] || 'Tổng tiền hàng';
    const targetCol2 = moneyColumns[1] || 'Tiền chiết khấu';
    const targetCol3 = moneyColumns[2] || moneyColumns[moneyColumns.length - 1] || 'Tổng tiền thanh toán';

    (duLieuDaLoc || []).forEach(item => {
      if (item && item.rowGoc) {
        const val1 = Number(String(item.rowGoc[targetCol1] || 0).replace(/,/g, '').trim()) || 0;
        const val2 = Number(String(item.rowGoc[targetCol2] || 0).replace(/,/g, '').trim()) || 0;
        const val3 = Number(String(item.rowGoc[targetCol3] || 0).replace(/,/g, '').trim()) || 0;

        col1Val += val1;
        col2Val += val2;
        col3Val += val3;
      }
    });

    return {
      name1: targetCol1.toUpperCase(),
      val1: col1Val,
      name2: targetCol2.toUpperCase(),
      val2: col2Val,
      name3: targetCol3.toUpperCase(),
      val3: col3Val
    };
  }, [duLieuDaLoc, tatCaTieuDeCotGoc]);

  const xuLySo = (val) => {
    if (!val) return 0;
    const clean = String(val).replace(/,/g, '').trim();
    return Number(clean) || 0;
  };

  const handleExportExcelDaSheet = () => {
    if (!duLieuHoaDon || duLieuHoaDon.length === 0) {
      alert("❌ Không có dữ liệu hóa đơn để xuất file!");
      return;
    }
   
    setLoadingExport(true);

    setTimeout(async () => {
      try {
        const workbook = new ExcelJS.Workbook();
        const thinBorder = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
        const baseFont = { name: 'Times New Roman', size: 12 };
        const boldFont = { name: 'Times New Roman', size: 12, bold: true };
        const titleFont = { name: 'Times New Roman', size: 13, bold: true };

        const nhomTheoCongTy = {};
        duLieuHoaDon.forEach(hd => {
          if (hd) {
            const cty = hd.tenCongTy ? hd.tenCongTy.trim() : 'Chưa rõ nhà cung cấp';
            if (!nhomTheoCongTy[cty]) nhomTheoCongTy[cty] = [];
            nhomTheoCongTy[cty].push(hd);
          }
        });

        const wsAll = workbook.addWorksheet('TỔNG HỢP TẤT CẢ');
        wsAll.mergeCells('B2:H2');
        wsAll.getCell('B2').value = 'DANH SÁCH HÓA ĐƠN TỔNG HỢP';
        wsAll.getCell('B2').font = titleFont;
        wsAll.getCell('B2').alignment = { horizontal: 'center', vertical: 'middle' };

        wsAll.addRow([]);
        const headers = ["", "STT", "SỐ HÓA ĐƠN", "NGÀY LẬP", "TÊN NGƯỜI BÁN", "TỔNG TIỀN HÀNG", "TIỀN CHIẾT KHẤU", "DOANH SỐ CHƯA THUẾ"];
        const hRow = wsAll.addRow(headers);
        hRow.height = 28;
        hRow.eachCell((cell, colNum) => {
          if (colNum >= 2) {
            cell.font = boldFont;
            cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
            cell.border = thinBorder;
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEAEAEA' } };
          }
        });

        duLieuHoaDon.forEach((item, idx) => {
          if (item && item.rowGoc) {
            const row = wsAll.addRow([
              "", idx + 1, 
              item.maHoaDon || "", 
              item.ngayHoaDon || "", 
              item.tenCongTy || "", 
              xuLySo(item.rowGoc["Tổng tiền hàng"] || 0), 
              xuLySo(item.rowGoc["Tiền chiết khấu"] || 0), 
              xuLySo(item.rowGoc["Doanh số bán chưa thuế"] || 0)
            ]);
            row.height = 22;
            row.eachCell((cell, colNum) => {
              if (colNum >= 2) {
                cell.font = baseFont;
                cell.border = thinBorder;
                if (colNum >= 6) {
                  cell.numFmt = '#,##0';
                  cell.alignment = { horizontal: 'right', vertical: 'middle' };
                }
              }
            });
          }
        });

        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const url = window.URL.createObjectURL(blob);
        const anchor = document.createElement('a');
        anchor.href = url;
        anchor.download = `BAO_CAO_HOA_DON_TONG_HOP.xlsx`;
        document.body.appendChild(anchor);
        anchor.click();
        document.body.removeChild(anchor);
        window.URL.revokeObjectURL(url);
      } catch (error) {
        console.error(error);
        alert("❌ Lỗi xuất file Excel!");
      } finally {
        setLoadingExport(false);
      }
    }, 100);
  };

  const getColumnWidth = (colName) => {
    if (!colName) return '140px';
    if (colName === 'STT') return '50px';
    if (colName.includes('Số hóa đơn') || colName.includes('Ngày')) return '110px';
    if (colName.includes('tiền') || colName.includes('Doanh số') || colName.includes('thuế')) return '140px';
    if (colName.includes('tên người bán') || colName.includes('Đơn vị')) return '280px';
    if (colName.includes('Diễn giải')) return '320px';
    return '140px';
  };

  return (
    <div style={{ background: 'white', padding: '24px', borderRadius: '8px', border: '1px solid #e2e8f0', boxShadow: '0 4px 6px rgba(0,0,0,0.05)' }}>
      
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '20px', borderBottom: '1px solid #e2e8f0', paddingBottom: '15px', flexWrap: 'wrap', gap: '15px' }}>
        <h3 style={{ color: '#805ad5', fontSize: '20px', fontWeight: 'bold', margin: 0 }}>
          DANH SÁCH HÓA ĐƠN CÔNG TY DƯỢC
        </h3>

        <div style={{ display: 'flex', gap: '10px', alignItems: 'center' }}>
          <label style={{ 
            background: '#3182ce', color: '#fff', padding: '9px 16px', borderRadius: '6px', cursor: 'pointer', fontWeight: '600', fontSize: '13px',
            boxShadow: '0 2px 4px rgba(49,130,206,0.2)', display: 'inline-flex', alignItems: 'center', gap: '6px'
          }}>
            📂 Nạp file Excel trực tiếp
            <input type="file" accept=".xlsx, .xls" onChange={handleDirectFileUpload} style={{ display: 'none' }} />
          </label>

          <button
            onClick={handleExportExcelDaSheet}
            disabled={loadingExport}
            style={{ padding: '9px 18px', backgroundColor: '#38a169', color: 'white', border: 'none', borderRadius: '6px', fontWeight: 'bold', fontSize: '14px', cursor: 'pointer', boxShadow: '0 2px 4px rgba(56, 161, 105, 0.2)', display: 'flex', alignItems: 'center', gap: '6px' }}
          >
            {loadingExport ? "⌛ Đang xuất file..." : "📊 Xuất file excel"}
          </button>
        </div>
      </div>
     
      <div style={{ marginBottom: '20px', padding: '15px', backgroundColor: '#f7fafc', borderRadius: '6px', border: '1px solid #e2e8f0' }}>
        <div style={{ display: 'flex', flexWrap: 'wrap', alignItems: 'center', justifyContent: 'space-between', gap: '20px', marginBottom: '15px' }}>
          <div style={{ display: 'flex', flexWrap: 'wrap', alignItems: 'center', gap: '15px' }}>
            <div>
              <span style={{ fontSize: '15px', fontWeight: '600', color: '#4a5568', marginRight: '10px' }}>Chọn đơn vị xem trên Web:</span>
              <select
                value={selectedCompany}
                onChange={(e) => setSelectedCompany(e.target.value)}
                style={{ padding: '8px 12px', borderRadius: '6px', border: '1px solid #cbd5e0', fontSize: '14px', outline: 'none', backgroundColor: '#fff', minWidth: '320px', fontWeight: '500' }}
              >
                {danhSachCongTy.map((cty, idx) => (
                  <option key={idx} value={cty}>{cty === 'ALL' ? '--- Xem tất cả công ty ---' : cty}</option>
                ))}
              </select>
            </div>
            <div style={{ fontSize: '14px', color: '#4a5568', paddingTop: '6px' }}>
              Hiển thị: <strong style={{ color: '#805ad5', fontSize: '16px' }}>{duLieuDaLoc.length}</strong> / {duLieuHoaDon?.length || 0} hóa đơn.
            </div>
          </div>
        </div>

        {/* 🌟 3 KHỐI THỐNG KÊ MỚI KHỚP TUYỆT ĐỐI VỚI 3 CỘT TRONG BẢNG */}
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: '15px', marginTop: '10px' }}>
          <div style={{ background: '#ebf8ff', padding: '12px', borderRadius: '6px', borderLeft: '4px solid #3182ce' }}>
            <div style={{ fontSize: '11px', color: '#2b6cb0', fontWeight: '700', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{thongKeTaiChinh.name1} (ĐANG XEM)</div>
            <div style={{ fontSize: '18px', fontWeight: '700', color: '#2b6cb0', marginTop: '4px' }}>{thongKeTaiChinh.val1.toLocaleString('vi-VN')} đ</div>
          </div>
          <div style={{ background: '#fffaf0', padding: '12px', borderRadius: '6px', borderLeft: '4px solid #dd6b20' }}>
            <div style={{ fontSize: '11px', color: '#dd6b20', fontWeight: '700', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{thongKeTaiChinh.name2} (ĐANG XEM)</div>
            <div style={{ fontSize: '18px', fontWeight: '700', color: '#dd6b20', marginTop: '4px' }}>{thongKeTaiChinh.val2.toLocaleString('vi-VN')} đ</div>
          </div>
          <div style={{ background: '#f0fff4', padding: '12px', borderRadius: '6px', borderLeft: '4px solid #38a169' }}>
            <div style={{ fontSize: '11px', color: '#38a169', fontWeight: '700', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{thongKeTaiChinh.name3} (ĐANG XEM)</div>
            <div style={{ fontSize: '18px', fontWeight: '700', color: '#38a169', marginTop: '4px' }}>{thongKeTaiChinh.val3.toLocaleString('vi-VN')} đ</div>
          </div>
        </div>
      </div>

      <div style={{ overflowX: 'auto', maxHeight: '550px', border: '1px solid #e2e8f0', borderRadius: '6px' }}>
        <table style={{ width: 'max-content', minWidth: '100%', borderCollapse: 'collapse', tableLayout: 'fixed', fontSize: '13px' }}>
          <thead style={{ position: 'sticky', top: 0, zIndex: 1 }}>
            <tr style={{ backgroundColor: '#edf2f7', color: '#2d3748', borderBottom: '2px solid #cbd5e0' }}>
              {tatCaTieuDeCotGoc.map((colName, index) => (
                <th
                  key={index}
                  style={{
                    width: getColumnWidth(colName), minWidth: getColumnWidth(colName), maxWidth: getColumnWidth(colName),
                    padding: '10px 8px', backgroundColor: '#edf2f7', fontWeight: '600',
                    textAlign: 'center', whiteSpace: 'normal', wordBreak: 'break-word', verticalAlign: 'middle',
                    borderRight: '1px solid #cbd5e0'
                  }}
                >
                  {colName}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {!duLieuDaLoc || duLieuDaLoc.length === 0 ? (
              <tr>
                <td colSpan={tatCaTieuDeCotGoc.length || 1} style={{ padding: '40px', textAlign: 'center', color: '#a0aec0', fontSize: '15px' }}>
                  📂 Chưa có dữ liệu hóa đơn. Vui lòng bấm nút <b>"Nạp file Excel trực tiếp"</b> ở trên để tải tệp!
                </td>
              </tr>
            ) : (
              duLieuDaLoc.map((item, index) => (
                <tr key={index} style={{ borderBottom: '1px solid #e2e8f0', backgroundColor: index % 2 === 0 ? '#fff' : '#f7fafc' }}>
                  {tatCaTieuDeCotGoc.map((colName, colIdx) => {
                    const value = item && item.rowGoc ? item.rowGoc[colName] : '';
                    const lk = colName.toLowerCase();
                    const isNumber = typeof value === 'number' || (lk.includes('tiền') || lk.includes('doanh số') || lk.includes('thuế') || lk.includes('thanh toán') || lk.includes('chiết khấu')) && !isNaN(value);
                    const widthStyle = getColumnWidth(colName);
                   
                    return (
                      <td
                        key={colIdx}
                        style={{
                          width: widthStyle, minWidth: widthStyle, maxWidth: widthStyle,
                          padding: '10px 8px', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis',
                          borderRight: '1px solid #e2e8f0',
                          textAlign: isNumber ? 'right' : (colName === 'STT' || colName.includes('Ngày') ? 'center' : 'left'),
                          fontWeight: colName.includes('Số hóa đơn') || lk.includes('thanh toán') ? '700' : 'normal',
                          color: lk.includes('thanh toán') ? '#2f855a' : '#2d3748'
                        }}
                        title={String(value || '')}
                      >
                        {isNumber && !isNaN(value) && value !== '' ? Number(value).toLocaleString('vi-VN') : String(value || '')}
                      </td>
                    );
                  })}
                </tr>
              ))
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
}

export default LocHoaDonCongTy;
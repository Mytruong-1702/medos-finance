import React, { useState, useMemo } from 'react';
import ExcelJS from 'exceljs';

function LocHoaDonCtyMST({ duLieuHoaDon }) {
  const [selectedCompany, setSelectedCompany] = useState('ALL');
  const [selectedMST, setSelectedMST] = useState('ALL');
  const [loadingExport, setLoadingExport] = useState(false);

  // Hàm hỗ trợ lấy tên công ty và loại bỏ hoàn toàn các dòng tiêu đề rác
  const layTenCongTy = (item) => {
    if (!item) return '';
    
    let rawName = '';
    if (item.tenCongTy && item.tenCongTy.trim() !== '') {
      rawName = item.tenCongTy.trim();
    } else if (item.rowGoc) {
      for (const key of Object.keys(item.rowGoc)) {
        if (key.toLowerCase().includes('tên người bán') || key.toLowerCase().includes('tên công ty')) {
          const val = item.rowGoc[key];
          if (val) {
            rawName = String(val).trim();
            break;
          }
        }
      }
    }

    const lower = rawName.toLowerCase();
    if (
      !rawName ||
      lower.includes('tên người bán/tên người xuất hàng') ||
      (lower.includes('tên người bán') && rawName.length < 15) ||
      lower.includes('tên công ty') ||
      lower.includes('stt') ||
      lower.includes('số hóa đơn')
    ) {
      return '';
    }

    return rawName;
  };

  // Hàm hỗ trợ lấy Mã số thuế
  const layMST = (item) => {
    if (!item) return '';
    if (item.mst && item.mst.trim() !== '') return item.mst.trim();
    if (item.rowGoc) {
      for (const key of Object.keys(item.rowGoc)) {
        if (key.toLowerCase().includes('mst') || key.toLowerCase().includes('mã số thuế')) {
          const val = item.rowGoc[key];
          if (val) return String(val).trim();
        }
      }
    }
    return '';
  };

  // Chuẩn hóa dữ liệu đầu vào, lọc bỏ dòng rác và SẮP XẾP GỘP THEO TỪNG CÔNG TY
  const danhSachGoc = useMemo(() => {
    let rawList = [];
    if (Array.isArray(duLieuHoaDon)) rawList = duLieuHoaDon;
    else if (duLieuHoaDon && Array.isArray(duLieuHoaDon.danhSach)) rawList = duLieuHoaDon.danhSach;
    
    const filtered = rawList.filter(item => {
      if (!item || !item.rowGoc) return false;
      const soHD = String(item.maHoaDon || item.rowGoc["Số hóa đơn"] || "").toLowerCase();
      const tenCty = String(item.tenCongTy || item.rowGoc["Tên người bán"] || "").toLowerCase();
      
      if (
        soHD.includes('số hóa đơn') || 
        soHD.includes('ký hiệu') || 
        soHD.includes('stt') ||
        tenCty.includes('tên người bán') ||
        tenCty.includes('tên công ty')
      ) {
        return false;
      }
      return true;
    });

    // Sắp xếp nhóm các hóa đơn lại với nhau theo Tên Công Ty (A -> Z)
    return filtered.sort((a, b) => {
      const ctyA = layTenCongTy(a).toLowerCase();
      const ctyB = layTenCongTy(b).toLowerCase();
      if (ctyA < ctyB) return -1;
      if (ctyA > ctyB) return 1;
      return 0;
    });
  }, [duLieuHoaDon]);

  const danhSachCongTy = useMemo(() => {
    if (danhSachGoc.length === 0) return ['ALL'];
    const ctySet = new Set();
    danhSachGoc.forEach(item => {
      const tenCty = layTenCongTy(item);
      if (tenCty) ctySet.add(tenCty);
    });
    return ['ALL', ...Array.from(ctySet)];
  }, [danhSachGoc]);

  const danhSachMST = useMemo(() => {
    if (danhSachGoc.length === 0) return ['ALL'];
    const mstSet = new Set();
    danhSachGoc.forEach(item => {
      const mst = layMST(item);
      if (mst) mstSet.add(mst);
    });
    return ['ALL', ...Array.from(mstSet)];
  }, [danhSachGoc]);

  const duLieuDaLoc = useMemo(() => {
    return danhSachGoc.filter(item => {
      if (!item) return false;
      const matchCty = selectedCompany === 'ALL' || layTenCongTy(item) === selectedCompany;
      const matchMST = selectedMST === 'ALL' || layMST(item) === selectedMST;
      return matchCty && matchMST;
    });
  }, [danhSachGoc, selectedCompany, selectedMST]);

  const tatCaTieuDeCotGoc = useMemo(() => {
    if (danhSachGoc.length === 0) return [];
    return Object.keys(danhSachGoc[0].rowGoc || {});
  }, [danhSachGoc]);

  const xuLySo = (val) => {
    if (val === undefined || val === null || val === '') return 0;
    if (typeof val === 'number') return val;
    const clean = String(val).replace(/,/g, '').trim();
    return Number(clean) || 0;
  };

  const layGiaTriSo = (rowGoc, keywords) => {
    if (!rowGoc) return 0;
    for (const kw of keywords) {
      for (const actualKey of Object.keys(rowGoc)) {
        if (actualKey.trim().toLowerCase().includes(kw.toLowerCase())) {
          const num = xuLySo(rowGoc[actualKey]);
          if (num !== 0) return num;
        }
      }
    }
    return 0;
  };

  const khoangThoiGianText = useMemo(() => {
    if (danhSachGoc.length === 0) return 'Từ ngày 01/09/2026 đến ngày 30/09/2026';
    
    for (const item of danhSachGoc) {
      if (item && item.rowGoc) {
        for (const key of Object.keys(item.rowGoc)) {
          const valStr = String(item.rowGoc[key] || '').trim();
          if (valStr.toLowerCase().includes('từ ngày') && valStr.toLowerCase().includes('đến ngày')) {
            return valStr;
          }
        }
      }
    }

    let dates = [];
    danhSachGoc.forEach(item => {
      if (item && item.rowGoc) {
        for (const key of Object.keys(item.rowGoc)) {
          if (key.toLowerCase().includes('ngày')) {
            const val = String(item.rowGoc[key] || '').trim();
            const parts = val.split('/');
            if (parts.length === 3) {
              const d = parseInt(parts[0], 10);
              const m = parseInt(parts[1], 10) - 1;
              const y = parseInt(parts[2], 10);
              const dateObj = new Date(y, m, d);
              if (!isNaN(dateObj.getTime())) {
                dates.push({ original: val, time: dateObj.getTime() });
              }
            }
          }
        }
      }
    });

    if (dates.length > 0) {
      dates.sort((a, b) => a.time - b.time);
      return `Từ ngày ${dates[0].original} đến ngày ${dates[dates.length - 1].original}`;
    }

    return 'Từ ngày 01/09/2026 đến ngày 30/09/2026';
  }, [danhSachGoc]);

  const thongKeTaiChinh = useMemo(() => {
    let tienHang = 0; let chietKhau = 0; let chuaThue = 0; let thanhToan = 0;
    duLieuDaLoc.forEach(item => {
      if (item && item.rowGoc) {
        tienHang += layGiaTriSo(item.rowGoc, ['tổng tiền hàng', 'tổng tiền']);
        chietKhau += layGiaTriSo(item.rowGoc, ['chiết khấu', 'tiền chiết']);
        chuaThue += layGiaTriSo(item.rowGoc, ['chưa thuế', 'doanh số']);
        thanhToan += layGiaTriSo(item.rowGoc, ['thanh toán']);
      }
    });
    return { tienHang, chietKhau, chuaThue, thanhToan };
  }, [duLieuDaLoc]);

  const handleExportExcelDaSheet = () => {
    if (danhSachGoc.length === 0) {
      alert("❌ Không có dữ liệu để xuất file!");
      return;
    }
    
    setLoadingExport(true);

    setTimeout(async () => {
      try {
        const workbook = new ExcelJS.Workbook();
        const thinBorder = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };

        const baseFont = { name: 'Times New Roman', size: 12 };
        const boldFont = { name: 'Times New Roman', size: 12, bold: true };
        const titleFont = { name: 'Times New Roman', size: 13, bold: true };

        const nhomTheoCongTy = {};
        danhSachGoc.forEach(hd => {
          if (hd) {
            const cty = layTenCongTy(hd);
            if (cty) {
              if (!nhomTheoCongTy[cty]) nhomTheoCongTy[cty] = [];
              nhomTheoCongTy[cty].push(hd);
            }
          }
        });

        // --- SHEET 1: TỔNG HỢP TẤT CẢ (SẮP XẾP GỘP CÔNG TY, CHỈ CÓ 1 DÒNG TỔNG CỘNG Ở CUỐI) ---
        const wsAll = workbook.addWorksheet('TỔNG HỢP TẤT CẢ');
        
        wsAll.mergeCells('B2:J2');
        const t1 = wsAll.getCell('B2');
        t1.value = 'DANH SÁCH HÓA ĐƠN TỔNG HỢP';
        t1.font = titleFont;
        t1.alignment = { horizontal: 'center', vertical: 'middle' };

        wsAll.mergeCells('B3:J3');
        const t2 = wsAll.getCell('B3');
        t2.value = khoangThoiGianText;
        t2.font = { name: 'Times New Roman', size: 11, italic: true };
        t2.alignment = { horizontal: 'center', vertical: 'middle' };

        wsAll.addRow([]);

        const headers = ["", "STT", "SỐ HÓA ĐƠN", "NGÀY LẬP", "TÊN NGƯỜI BÁN", "MÃ SỐ THUẾ", "TỔNG TIỀN HÀNG", "TIỀN CHIẾT KHẤU", "DOANH SỐ CHƯA THUẾ", "TỔNG THANH TOÁN"];
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

        let sumTienHangAll = 0, sumChietKhauAll = 0, sumChuaThueAll = 0, sumThanhToanAll = 0;
        let sttGlobal = 1;

        Object.keys(nhomTheoCongTy).forEach(tenCty => {
          const dsHoaDonCty = nhomTheoCongTy[tenCty];
          
          dsHoaDonCty.forEach((item) => {
            if (item && item.rowGoc) {
              const tienHang = layGiaTriSo(item.rowGoc, ['tổng tiền hàng', 'tổng tiền']);
              const chietKhau = layGiaTriSo(item.rowGoc, ['chiết khấu', 'tiền chiết']);
              const chuaThue = layGiaTriSo(item.rowGoc, ['chưa thuế', 'doanh số']);
              const thanhToan = layGiaTriSo(item.rowGoc, ['thanh toán']);

              sumTienHangAll += tienHang;
              sumChietKhauAll += chietKhau;
              sumChuaThueAll += chuaThue;
              sumThanhToanAll += thanhToan;

              const soHD = String(item.maHoaDon || item.rowGoc["Số hóa đơn"] || "").trim();
              const ngayHD = String(item.ngayHoaDon || item.rowGoc["Ngày hóa đơn"] || item.rowGoc["Ngày lập"] || "").trim();
              const mstCty = layMST(item);

              const row = wsAll.addRow(["", sttGlobal++, soHD, ngayHD, tenCty, mstCty, tienHang, chietKhau, chuaThue, thanhToan]);
              row.height = 22;
              row.eachCell((cell, colNum) => {
                if (colNum >= 2) {
                  cell.font = baseFont;
                  cell.border = thinBorder;
                  if (colNum === 2 || colNum === 4 || colNum === 6) cell.alignment = { horizontal: 'center', vertical: 'middle' };
                  if (colNum === 3) cell.alignment = { horizontal: 'center', vertical: 'middle' };
                  if (colNum === 5) cell.alignment = { horizontal: 'left', vertical: 'middle' };
                  if (colNum >= 7) {
                    cell.numFmt = '#,##0';
                    cell.alignment = { horizontal: 'right', vertical: 'middle' };
                  }
                }
              });
            }
          });
        });

        const totRow = wsAll.addRow(["", "", "Tổng cộng", "", "", "", sumTienHangAll, sumChietKhauAll, sumChuaThueAll, sumThanhToanAll]);
        totRow.height = 24;
        totRow.eachCell((cell, colNum) => {
          if (colNum >= 2) {
            cell.border = thinBorder;
            cell.font = boldFont;
            if (colNum >= 7) {
              cell.numFmt = '#,##0';
              cell.alignment = { horizontal: 'right', vertical: 'middle' };
            }
          }
        });
        wsAll.columns = [{ width: 2 }, { width: 6 }, { width: 16 }, { width: 14 }, { width: 38 }, { width: 16 }, { width: 18 }, { width: 16 }, { width: 20 }, { width: 22 }];

        // --- CÁC SHEET TIẾP THEO: TÁCH RIÊNG TỪNG CÔNG TY ---
        const danhSachCtyThucTe = Object.keys(nhomTheoCongTy);
        const danhSachTenSheetDaDung = new Set();

        danhSachCtyThucTe.forEach((tenCty, index) => {
          const hoaDonCuaCty = nhomTheoCongTy[tenCty] || [];
          if (hoaDonCuaCty.length > 0) {
            let sheetTitle = tenCty.replace(/CÔNG TY CỔ PHẦN/g, 'CTCP').replace(/CHI NHÁNH/g, 'CN');
            let tenSạch = sheetTitle.replace(/[\\\/ \?\*\:\[\]]/g, "_").substring(0, 30);
            
            if (danhSachTenSheetDaDung.has(tenSạch)) {
              tenSạch = `${index + 1}_${Math.floor(Math.random() * 100)}`;
            }
            danhSachTenSheetDaDung.add(tenSạch);

            const wsCty = workbook.addWorksheet(tenSạch);

            wsCty.mergeCells('B2:D2');
            wsCty.mergeCells('E2:G2');
            wsCty.getCell('B2').value = 'SỞ Y TẾ TỈNH CÀ MAU';
            wsCty.getCell('B2').font = boldFont;
            wsCty.getCell('E2').value = 'CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM';
            wsCty.getCell('E2').font = boldFont;
            wsCty.getCell('E2').alignment = { horizontal: 'center', vertical: 'middle' };

            wsCty.mergeCells('B3:D3');
            wsCty.mergeCells('E3:G3');
            wsCty.getCell('B3').value = 'CTCP PKĐK Y ĐỨC SÀI GÒN';
            wsCty.getCell('B3').font = baseFont;
            wsCty.getCell('E3').value = 'Độc lập - Tự do - Hạnh phúc';
            wsCty.getCell('E3').font = baseFont;
            wsCty.getCell('E3').alignment = { horizontal: 'center', vertical: 'middle' };

            wsCty.addRow([]);
            
            wsCty.mergeCells('B5:G5');
            const titleCellCty = wsCty.getCell('B5');
            titleCellCty.value = tenCty.toUpperCase();
            titleCellCty.font = titleFont;
            titleCellCty.alignment = { horizontal: 'center', vertical: 'middle' };

            wsCty.mergeCells('B6:G6');
            const subTitleCty = wsCty.getCell('B6');
            subTitleCty.value = khoangThoiGianText;
            subTitleCty.font = { name: 'Times New Roman', size: 11, italic: true };
            subTitleCty.alignment = { horizontal: 'center', vertical: 'middle' };

            wsCty.addRow([]);

            const hRowC = wsCty.addRow(["", "STT", "SỐ HÓA ĐƠN", "NGÀY LẬP", "TỔNG TIỀN HÀNG", "DOANH SỐ CHƯA THUẾ", "TỔNG THANH TOÁN"]);
            hRowC.height = 28;
            hRowC.eachCell((cell, colNum) => {
              if (colNum >= 2) {
                cell.font = boldFont;
                cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
                cell.border = thinBorder;
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } };
              }
            });

            let cSumTienHang = 0, cSumChuaThue = 0, cSumThanhToan = 0;

            hoaDonCuaCty.forEach((item, idx) => {
              if (item && item.rowGoc) {
                const tienHang = layGiaTriSo(item.rowGoc, ['tổng tiền hàng', 'tổng tiền']);
                const chuaThue = layGiaTriSo(item.rowGoc, ['chưa thuế', 'doanh số']);
                const thanhToan = layGiaTriSo(item.rowGoc, ['thanh toán']);

                cSumTienHang += tienHang;
                cSumChuaThue += chuaThue;
                cSumThanhToan += thanhToan;

                const soHD = String(item.maHoaDon || item.rowGoc["Số hóa đơn"] || "").trim();
                const ngayHD = String(item.ngayHoaDon || item.rowGoc["Ngày hóa đơn"] || item.rowGoc["Ngày lập"] || "").trim();

                const rowC = wsCty.addRow(["", idx + 1, soHD, ngayHD, tienHang, chuaThue, thanhToan]);
                rowC.height = 22;
                rowC.eachCell((cell, colNum) => {
                  if (colNum >= 2) {
                    cell.border = thinBorder;
                    if (colNum === 2 || colNum === 4) cell.alignment = { horizontal: 'center', vertical: 'middle' };
                    if (colNum === 3) cell.alignment = { horizontal: 'center', vertical: 'middle' };
                    if (colNum >= 5) {
                      cell.numFmt = '#,##0';
                      cell.alignment = { horizontal: 'right', vertical: 'middle' };
                    }
                  }
                });
              }
            });

            const totRowC = wsCty.addRow(["", "", "Tổng cộng", "", cSumTienHang, cSumChuaThue, cSumThanhToan]);
            totRowC.height = 24;
            totRowC.eachCell((cell, colNum) => {
              if (colNum >= 2) {
                cell.border = thinBorder;
                cell.font = boldFont;
                if (colNum >= 5) {
                  cell.numFmt = '#,##0';
                  cell.alignment = { horizontal: 'right', vertical: 'middle' };
                }
              }
            });

            wsCty.addRow([]);
            wsCty.mergeCells(`E${wsCty.rowCount + 1}:G${wsCty.rowCount + 1}`);
            wsCty.getCell(`E${wsCty.rowCount}`).value = 'Cà Mau, ngày ... tháng ... năm ...';
            wsCty.getCell(`E${wsCty.rowCount}`).font = { name: 'Times New Roman', size: 12, italic: true };
            wsCty.getCell(`E${wsCty.rowCount}`).alignment = { horizontal: 'center' };

            const rSignNum = wsCty.rowCount + 1;
            const rSign = wsCty.addRow([]);
            rSign.height = 55;
            
            wsCty.mergeCells(`B${rSignNum}:D${rSignNum}`);
            wsCty.mergeCells(`E${rSignNum}:G${rSignNum}`);
            
            const cellLeft = wsCty.getCell(`B${rSignNum}`);
            cellLeft.value = {
              richText: [
                { text: 'Người lập biểu\n', font: { name: 'Times New Roman', size: 12, bold: true } },
                { text: '(Ký ghi rõ họ tên)', font: { name: 'Times New Roman', size: 11, italic: true, bold: false } }
              ]
            };
            cellLeft.alignment = { horizontal: 'center', vertical: 'top', wrapText: true };

            const cellRight = wsCty.getCell(`E${rSignNum}`);
            cellRight.value = {
              richText: [
                { text: 'Kế toán\n', font: { name: 'Times New Roman', size: 12, bold: true } },
                { text: '(Ký ghi rõ họ tên)', font: { name: 'Times New Roman', size: 11, italic: true, bold: false } }
              ]
            };
            cellRight.alignment = { horizontal: 'center', vertical: 'top', wrapText: true };

            wsCty.columns = [{ width: 2 }, { width: 6 }, { width: 16 }, { width: 14 }, { width: 18 }, { width: 20 }, { width: 22 }];
          }
        });

        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const url = window.URL.createObjectURL(blob);
        const anchor = document.createElement('a');
        anchor.href = url;
        anchor.download = `BAO_CAO_HOA_DON_CTY_MST.xlsx`;
        document.body.appendChild(anchor);
        anchor.click();
        document.body.removeChild(anchor);
        window.URL.revokeObjectURL(url);
        
      } catch (error) {
        console.error(error);
        alert("❌ Có lỗi phát sinh khi xuất file Excel!");
      } finally {
        setLoadingExport(false);
      }
    }, 100);
  };

  const getColumnWidth = (colName) => {
    if (colName === 'STT') return '50px';
    if (colName === 'Số hóa đơn' || colName === 'Ngày hóa đơn' || colName === 'Ngày lập') return '110px';
    if (colName.includes('MST') || colName.includes('Mã số thuế')) return '140px';
    if (colName.includes('Tổng tiền') || colName.includes('Doanh số') || colName.includes('Tiền chiết khấu')) return '130px'; 
    if (colName.includes('Tên người bán')) return '320px';
    if (colName.includes('Diễn giải')) return '380px';
    return '140px';
  };

  return (
    <div style={{ background: 'white', padding: '24px', borderRadius: '8px', border: '1px solid #e2e8f0', boxShadow: '0 4px 6px rgba(0,0,0,0.05)' }}>
      <h3 style={{ color: '#805ad5', fontSize: '20px', fontWeight: 'bold', marginBottom: '15px' }}>
        LỌC HÓA ĐƠN CTY & MST
      </h3>
      
      <div style={{ marginBottom: '20px', padding: '15px', backgroundColor: '#f7fafc', borderRadius: '6px', border: '1px solid #e2e8f0' }}>
        <div style={{ display: 'flex', flexWrap: 'wrap', alignItems: 'center', justifyContent: 'space-between', gap: '20px', marginBottom: '15px' }}>
          <div style={{ display: 'flex', flexWrap: 'wrap', alignItems: 'center', gap: '15px' }}>
            <div>
              <span style={{ fontSize: '14px', fontWeight: '600', color: '#4a5568', display: 'block', marginBottom: '4px' }}>Chọn đơn vị (Công ty):</span>
              <select
                value={selectedCompany}
                onChange={(e) => setSelectedCompany(e.target.value)}
                style={{ padding: '8px 12px', borderRadius: '6px', border: '1px solid #cbd5e0', fontSize: '14px', outline: 'none', backgroundColor: '#fff', minWidth: '280px', fontWeight: '500' }}
              >
                {danhSachCongTy.map((cty, idx) => (
                  <option key={idx} value={cty}>{cty === 'ALL' ? '--- Tất cả công ty ---' : cty}</option>
                ))}
              </select>
            </div>

            <div>
              <span style={{ fontSize: '14px', fontWeight: '600', color: '#4a5568', display: 'block', marginBottom: '4px' }}>Lọc theo Mã số thuế (MST):</span>
              <select
                value={selectedMST}
                onChange={(e) => setSelectedMST(e.target.value)}
                style={{ padding: '8px 12px', borderRadius: '6px', border: '1px solid #cbd5e0', fontSize: '14px', outline: 'none', backgroundColor: '#fff', minWidth: '180px', fontWeight: '500' }}
              >
                {danhSachMST.map((mst, idx) => (
                  <option key={idx} value={mst}>{mst === 'ALL' ? '--- Tất cả MST ---' : mst}</option>
                ))}
              </select>
            </div>

            <div style={{ fontSize: '14px', color: '#4a5568', paddingTop: '22px' }}>
              Đang hiển thị: <strong style={{ color: '#805ad5', fontSize: '16px' }}>{duLieuDaLoc.length}</strong> / {danhSachGoc.length} hóa đơn.
            </div>
          </div>

          <button
            onClick={handleExportExcelDaSheet}
            disabled={loadingExport}
            style={{ padding: '9px 18px', backgroundColor: '#38a169', color: 'white', border: 'none', borderRadius: '6px', fontWeight: 'bold', fontSize: '14px', cursor: 'pointer', boxShadow: '0 2px 4px rgba(56, 161, 105, 0.2)', display: 'flex', alignItems: 'center', gap: '6px', marginTop: '22px', transition: '0.2s' }}
            onMouseOver={(e) => e.target.style.backgroundColor = '#2f855a'}
            onMouseOut={(e) => e.target.style.backgroundColor = '#38a169'}
          >
            {loadingExport ? "⌛ Đang xuất file..." : "📊 Xuất file excel"}
          </button>
        </div>

        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr 1fr', gap: '15px', marginTop: '10px' }}>
          <div style={{ background: '#ebf8ff', padding: '12px', borderRadius: '6px', borderLeft: '4px solid #3182ce' }}>
            <div style={{ fontSize: '11px', color: '#2b6cb0', fontWeight: '600' }}>TỔNG TIỀN HÀNG</div>
            <div style={{ fontSize: '16px', fontWeight: '700', color: '#2b6cb0' }}>{thongKeTaiChinh.tienHang.toLocaleString('vi-VN')} đ</div>
          </div>
          <div style={{ background: '#fffaf0', padding: '12px', borderRadius: '6px', borderLeft: '4px solid #dd6b20' }}>
            <div style={{ fontSize: '11px', color: '#dd6b20', fontWeight: '600' }}>TIỀN CHIẾT KHẤU</div>
            <div style={{ fontSize: '16px', fontWeight: '700', color: '#dd6b20' }}>{thongKeTaiChinh.chietKhau.toLocaleString('vi-VN')} đ</div>
          </div>
          <div style={{ background: '#f5f3ff', padding: '12px', borderRadius: '6px', borderLeft: '4px solid #7c3aed' }}>
            <div style={{ fontSize: '11px', color: '#7c3aed', fontWeight: '600' }}>DOANH SỐ CHƯA THUẾ</div>
            <div style={{ fontSize: '16px', fontWeight: '700', color: '#7c3aed' }}>{thongKeTaiChinh.chuaThue.toLocaleString('vi-VN')} đ</div>
          </div>
          <div style={{ background: '#f0fff4', padding: '12px', borderRadius: '6px', borderLeft: '4px solid #38a169' }}>
            <div style={{ fontSize: '11px', color: '#38a169', fontWeight: '600' }}>TỔNG THANH TOÁN</div>
            <div style={{ fontSize: '16px', fontWeight: '700', color: '#38a169' }}>{thongKeTaiChinh.thanhToan.toLocaleString('vi-VN')} đ</div>
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
            {duLieuDaLoc.length === 0 ? (
              <tr>
                <td colSpan={tatCaTieuDeCotGoc.length || 1} style={{ padding: '20px', textAlign: 'center', color: '#a0aec0' }}>Không có dữ liệu phù hợp.</td>
              </tr>
            ) : (
              duLieuDaLoc.map((item, index) => (
                <tr key={index} style={{ borderBottom: '1px solid #e2e8f0', backgroundColor: index % 2 === 0 ? '#fff' : '#f7fafc' }}>
                  {tatCaTieuDeCotGoc.map((colName, colIdx) => {
                    const value = item && item.rowGoc ? item.rowGoc[colName] : '';
                    const isNumber = typeof value === 'number' || (colName.includes('Tổng tiền') || colName.includes('Doanh số') || colName.includes('Tiền chiết khấu')) && !isNaN(value);
                    const widthStyle = getColumnWidth(colName);
                   
                    return (
                      <td
                        key={colIdx}
                        style={{
                          width: widthStyle, minWidth: widthStyle, maxWidth: widthStyle,
                          padding: '10px 8px', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis',
                          borderRight: '1px solid #e2e8f0',
                          textAlign: isNumber ? 'right' : (colName === 'STT' || colName === 'Ngày lập' || colName === 'Ngày hóa đơn' || colName === 'MST' ? 'center' : 'left'),
                          fontWeight: colName === 'Số hóa đơn' || colName === 'Tổng tiền thanh toán' ? '700' : 'normal',
                          color: colName === 'Tổng tiền thanh toán' ? '#2f855a' : '#2d3748'
                        }}
                        title={String(value || '')}
                      >
                        {isNumber ? Number(value).toLocaleString('vi-VN') : String(value || '')}
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

export default LocHoaDonCtyMST;
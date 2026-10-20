import React, { useState, useMemo } from 'react';
import ExcelJS from 'exceljs';

function TongHopYLenhBacSi({ duLieuTongHop = [], duLieuTongHopPhong = [] }) {
  const [cheDoXem, setCheDoXem] = useState('BAC_SI');
  const [selectedDoctor, setSelectedDoctor] = useState('ALL');
  const [selectedPhong, setSelectedPhong] = useState('ALL');
  const [ngayThangBaoCao, setNgayThangBaoCao] = useState('Từ ngày 01-09-2026 đến ngày 30-09-2026');

  const formatNgayDMY = (text) => {
    if (!text) return text;
    return String(text).replace(/(\d{4})-(\d{2})-(\d{2})/g, '$3-$2-$1')
                       .replace(/(\d{4})\/(\d{2})\/(\d{2})/g, '$3-$2-$1');
  };

  const chuanHoachuoi = (str) => {
    if (!str) return '';
    let s = String(str).trim();
    if (/[\u0080-\u00FF]{4,}/.test(s) || s.includes('Ð') || s.includes('Öiw')) {
      return `Dịch vụ y tế (${s.substring(0, 8)}...)`;
    }
    return s.replace(/[\u0000-\u001F\u007F-\u009F]/g, '');
  };

  const chuanHoaTenNhom = (text) => {
    const upper = String(text || '').toUpperCase();
    if (upper.includes('CDHA') || upper.includes('CHẨN ĐOÁN HÌNH ẢNH')) return 'CDHA';
    if (upper.includes('CÔNG KHÁM') || upper.includes('KHÁM')) return 'Công khám';
    if (upper.includes('THUỐC') || upper.includes('VTYT') || upper.includes('VẬT TƯ')) return 'Thuốc và VTYT';
    if (upper.includes('XÉT NGHIỆM')) return 'Xét nghiệm';
    if (upper.includes('THỦ THUẬT') || upper.includes('PHẪU THUẬT') || upper.includes('VLTL')) return 'Thủ thuật - Phẫu thuật - VLTL';
    return 'Thủ thuật - Phẫu thuật - VLTL';
  };

  const parseDoctorDataToRows = (rawArray) => {
    if (!rawArray || !Array.isArray(rawArray) || rawArray.length === 0) return [];
    
    const normalizedRows = rawArray.map(item => {
      if (Array.isArray(item)) return item;
      if (item && typeof item === 'object') return Object.values(item);
      return [];
    });

    let targetRows = [];
    let currentDoc = "Bác sĩ tổng hợp";
    let currentNhomChuyenKhoa = "CDHA";
    let startParsing = false;
    let detectedDate = '';
    
    let colIdx = { stt: 1, tenDV: 2, loai: 3, soLuong: 4, donGia: 5, thanhTien: 6 };

    for (let i = 0; i < normalizedRows.length; i++) {
      const row = normalizedRows[i];
      if (!row || !Array.isArray(row) || row.length === 0) continue;

      if (!detectedDate && i < 15) {
        const rowText = row.map(c => c ? String(c).trim() : '').join(' ');
        if (rowText.toLowerCase().includes('từ ngày') || rowText.toLowerCase().includes('ngày')) {
          detectedDate = rowText;
        }
      }

      const rowStr = row.map(c => String(c || '').toUpperCase());
      
      if (!startParsing) {
        if (rowStr.some(c => c.includes('TÊN DỊCH VỤ') || c.includes('TEN DICH VU') || c.includes('THANH TIỀN'))) {
          startParsing = true;
          const idxDV = rowStr.findIndex(c => c.includes('TÊN DỊCH VỤ') || c.includes('TEN DICH VU'));
          if (idxDV !== -1) colIdx.tenDV = idxDV;
          const idxLoai = rowStr.findIndex(c => c.includes('LOẠI') || c.includes('LOAI'));
          if (idxLoai !== -1) colIdx.loai = idxLoai;
          const idxSL = rowStr.findIndex(c => c.includes('SỐ LƯỢNG') || c.includes('SO LUONG') || c.includes('SL'));
          if (idxSL !== -1) colIdx.soLuong = idxSL;
          const idxDG = rowStr.findIndex(c => c.includes('ĐƠN GIÁ') || c.includes('DON GIA'));
          if (idxDG !== -1) colIdx.donGia = idxDG;
          const idxTT = rowStr.findIndex(c => c.includes('THÀNH TIỀN') || c.includes('THANH TIEN'));
          if (idxTT !== -1) colIdx.thanhTien = idxTT;
          continue;
        }
      }

      if (startParsing) {
        const rowJoinedText = row.map(c => c ? String(c).trim() : '').join(' ').toUpperCase();
        
        if (rowJoinedText.replace(/\s+/g, '') === '' || rowJoinedText.includes('TRANG')) continue;
        if (rowJoinedText.includes('SỞ Y TẾ') || rowJoinedText.includes('PKĐK') || rowJoinedText.includes('BÁO CÁO DOANH THU')) continue;

        const parseNum = (val) => {
          if (val === undefined || val === null || val === '' ) return 0;
          const cleanStr = String(val).replace(/\./g, '').replace(/,/g, '.');
          return Number(cleanStr) || 0;
        };

        const fullRowText = row.map(c => c ? String(c).trim() : '').filter(Boolean).join(' ');

        const isDoctorRow = (fullRowText.toUpperCase().includes('BS') || fullRowText.toUpperCase().includes('BÁC SĨ') || fullRowText.toUpperCase().includes('BS.')) &&
                            !fullRowText.toUpperCase().includes('TỔNG CỘNG') &&
                            parseNum(row[4]) === 0 && parseNum(row[6]) === 0 &&
                            !fullRowText.toUpperCase().includes('TÊN DỊCH VỤ');

        if (isDoctorRow) {
          currentDoc = fullRowText;
          continue;
        }

        const cellB = row[1] ? String(row[1]).trim() : '';
        const cellC = row[2] ? String(row[2]).trim() : '';

        const isGroupHeaderRow = (cellB && !Number(cellB) && !cellC && 
                                  !cellB.toUpperCase().includes('TỔNG CỘNG') && 
                                  parseNum(row[4]) === 0 && parseNum(row[6]) === 0) ||
                                 (cellB && (cellB.toUpperCase() === 'CDHA' || cellB.toUpperCase() === 'XÉT NGHIỆM' || cellB.toUpperCase() === 'CÔNG KHÁM' || cellB.toUpperCase().includes('THUỐC')));

        if (isGroupHeaderRow && !cellB.toUpperCase().includes('BS') && !cellB.toUpperCase().includes('BÁC SĨ')) {
          currentNhomChuyenKhoa = chuanHoaTenNhom(cellB);
          continue;
        }

        if (rowJoinedText.includes('TỔNG CỘNG') || rowJoinedText.includes('TONG CONG')) {
          continue;
        }

        const tenDV = row[colIdx.tenDV] || row[2];
        const cleanTenDV = chuanHoachuoi(tenDV);
        const soLuong = row[colIdx.soLuong] || row[4];
        const donGia = row[colIdx.donGia] || row[5];
        const thanhTien = row[colIdx.thanhTien] || row[6];

        const upperDV = cleanTenDV.toUpperCase();
        const laTieuDeThua = upperDV === 'CDHA' || upperDV === 'XÉT NGHIỆM' || upperDV === 'THUỐC' || upperDV === 'CÔNG KHÁM' || upperDV === 'THỦ THUẬT';

        if (cleanTenDV && upperDV !== 'TÊN DỊCH VỤ' && cleanTenDV.length > 1 && !laTieuDeThua) {
          targetRows.push({
            tenNhomDoiTuong: currentDoc,
            tenNhomPhanLoai: currentNhomChuyenKhoa,
            tenDichVu: cleanTenDV,
            loai: (row[colIdx.loai] ? chuanHoachuoi(row[colIdx.loai]) : 'BHYT'),
            soLuong: parseNum(soLuong),
            donGia: parseNum(donGia),
            thanhTien: parseNum(thanhTien)
          });
        }
      }
    }

    if (detectedDate) setNgayThangBaoCao(formatNgayDMY(detectedDate));
    return targetRows;
  };

  const parsePhongDataToRows = (rawArray) => {
    if (!rawArray || !Array.isArray(rawArray) || rawArray.length === 0) return [];
    
    const normalizedRows = rawArray.map(item => {
      if (Array.isArray(item)) return item;
      if (item && typeof item === 'object') return Object.values(item);
      return [];
    });

    let targetRows = [];
    let currentPhongName = "Phòng Mặc Định";
    let currentNhomChuyenKhoa = "CDHA";
    let startParsing = false;
    let detectedDate = '';
    
    let colIdx = { stt: 1, tenDV: 2, loai: 3, soLuong: 4, donGia: 5, thanhTien: 6 };

    for (let i = 0; i < normalizedRows.length; i++) {
      const row = normalizedRows[i];
      if (!row || !Array.isArray(row) || row.length === 0) continue;

      if (!detectedDate && i < 15) {
        const rowText = row.map(c => c ? String(c).trim() : '').join(' ');
        if (rowText.toLowerCase().includes('từ ngày') || rowText.toLowerCase().includes('ngày')) {
          detectedDate = rowText;
        }
      }

      const rowStr = row.map(c => String(c || '').toUpperCase());
      
      if (!startParsing) {
        if (rowStr.some(c => c.includes('TÊN DỊCH VỤ') || c.includes('TEN DICH VU') || c.includes('THANH TIỀN'))) {
          startParsing = true;
          const idxDV = rowStr.findIndex(c => c.includes('TÊN DỊCH VỤ') || c.includes('TEN DICH VU'));
          if (idxDV !== -1) colIdx.tenDV = idxDV;
          const idxLoai = rowStr.findIndex(c => c.includes('LOẠI') || c.includes('LOAI'));
          if (idxLoai !== -1) colIdx.loai = idxLoai;
          const idxSL = rowStr.findIndex(c => c.includes('SỐ LƯỢNG') || c.includes('SO LUONG') || c.includes('SL'));
          if (idxSL !== -1) colIdx.soLuong = idxSL;
          const idxDG = rowStr.findIndex(c => c.includes('ĐƠN GIÁ') || c.includes('DON GIA'));
          if (idxDG !== -1) colIdx.donGia = idxDG;
          const idxTT = rowStr.findIndex(c => c.includes('THÀNH TIỀN') || c.includes('THANH TIEN'));
          if (idxTT !== -1) colIdx.thanhTien = idxTT;
          continue;
        }
      }

      if (startParsing) {
        const rowJoinedText = row.map(c => c ? String(c).trim() : '').join(' ').toUpperCase();
        
        if (rowJoinedText.replace(/\s+/g, '') === '' || rowJoinedText.includes('TRANG')) continue;
        if (rowJoinedText.includes('SỞ Y TẾ') || rowJoinedText.includes('PKĐK') || rowJoinedText.includes('BÁO CÁO DOANH THU')) continue;

        const parseNum = (val) => {
          if (val === undefined || val === null || val === '' ) return 0;
          const cleanStr = String(val).replace(/\./g, '').replace(/,/g, '.');
          return Number(cleanStr) || 0;
        };

        const fullRowText = row.map(c => c ? String(c).trim() : '').filter(Boolean).join(' ');
        
        const isPhongRow = (fullRowText.toUpperCase().includes('PHÒNG') || fullRowText.toUpperCase().includes('KHOA')) &&
                           !fullRowText.toUpperCase().includes('TỔNG CỘNG') &&
                           parseNum(row[4]) === 0 && parseNum(row[6]) === 0 &&
                           !fullRowText.toUpperCase().includes('TÊN DỊCH VỤ');

        if (isPhongRow) {
          currentPhongName = fullRowText;
          continue;
        }

        const cellB = row[1] ? String(row[1]).trim() : '';
        const cellC = row[2] ? String(row[2]).trim() : '';

        const isGroupHeaderRow = (cellB && !Number(cellB) && !cellC && 
                                  !cellB.toUpperCase().includes('TỔNG CỘNG') && 
                                  parseNum(row[4]) === 0 && parseNum(row[6]) === 0) ||
                                 (cellB && (cellB.toUpperCase() === 'CDHA' || cellB.toUpperCase() === 'XÉT NGHIỆM' || cellB.toUpperCase() === 'CÔNG KHÁM' || cellB.toUpperCase().includes('THUỐC')));

        if (isGroupHeaderRow && !cellB.toUpperCase().includes('PHÒNG') && !cellB.toUpperCase().includes('KHOA')) {
          currentNhomChuyenKhoa = chuanHoaTenNhom(cellB);
          continue;
        }

        if (rowJoinedText.includes('TỔNG CỘNG') || rowJoinedText.includes('TONG CONG')) {
          continue;
        }

        const tenDV = row[colIdx.tenDV] || row[2];
        const cleanTenDV = chuanHoachuoi(tenDV);
        const soLuong = row[colIdx.soLuong] || row[4];
        const donGia = row[colIdx.donGia] || row[5];
        const thanhTien = row[colIdx.thanhTien] || row[6];

        const upperDV = cleanTenDV.toUpperCase();
        const laTieuDeThua = upperDV === 'CDHA' || upperDV === 'XÉT NGHIỆM' || upperDV === 'THUỐC' || upperDV === 'CÔNG KHÁM' || upperDV === 'THỦ THUẬT';

        if (cleanTenDV && upperDV !== 'TÊN DỊCH VỤ' && cleanTenDV.length > 1 && !laTieuDeThua) {
          targetRows.push({
            tenNhomDoiTuong: currentPhongName,
            tenNhomPhanLoai: currentNhomChuyenKhoa,
            tenDichVu: cleanTenDV,
            loai: (row[colIdx.loai] ? chuanHoachuoi(row[colIdx.loai]) : 'BHYT'),
            soLuong: parseNum(soLuong),
            donGia: parseNum(donGia),
            thanhTien: parseNum(thanhTien)
          });
        }
      }
    }

    if (detectedDate) setNgayThangBaoCao(formatNgayDMY(detectedDate));
    return targetRows;
  };

  const xuLyDuLieuBacSi = useMemo(() => parseDoctorDataToRows(duLieuTongHop), [duLieuTongHop]);
  const xuLyDuLieuPhong = useMemo(() => {
    const sourceData = (duLieuTongHopPhong && duLieuTongHopPhong.length > 0) ? duLieuTongHopPhong : duLieuTongHop;
    return parsePhongDataToRows(sourceData).filter(item => item.tenNhomDoiTuong && item.tenNhomDoiTuong.toLowerCase() !== 'null' && item.tenNhomDoiTuong !== '');
  }, [duLieuTongHopPhong, duLieuTongHop]);

  // Giữ nguyên từng dòng chi tiết (không gom nhóm cộng dồn) để hiển thị đầy đủ các mức giá khác nhau
  const taoCauTrucChiTiet = (rawRows) => {
    const mapDoiTuong = {};
    rawRows.forEach(item => {
      const { tenNhomDoiTuong, tenNhomPhanLoai, tenDichVu, loai, soLuong, donGia, thanhTien } = item;
      const nhomChuan = tenNhomPhanLoai || 'CDHA';

      if (!mapDoiTuong[tenNhomDoiTuong]) {
        mapDoiTuong[tenNhomDoiTuong] = { 
          ten: tenNhomDoiTuong, 
          cacNhóm: { 
            'Công khám': [], 
            'CDHA': [], 
            'Xét nghiệm': [], 
            'Thuốc và VTYT': [],
            'Thủ thuật - Phẫu thuật - VLTL': [] 
          } 
        };
      }
      if (!mapDoiTuong[tenNhomDoiTuong].cacNhóm[nhomChuan]) {
        mapDoiTuong[tenNhomDoiTuong].cacNhóm[nhomChuan] = [];
      }
      mapDoiTuong[tenNhomDoiTuong].cacNhóm[nhomChuan].push({ tenDichVu, loai, soLuong, donGia, thanhTien });
    });
    return Object.values(mapDoiTuong);
  };

  const cauTrucTheoBacSi = useMemo(() => taoCauTrucChiTiet(xuLyDuLieuBacSi), [xuLyDuLieuBacSi]);
  const cauTrucTheoPhong = useMemo(() => taoCauTrucChiTiet(xuLyDuLieuPhong), [xuLyDuLieuPhong]);

  const danhSachTenBS = useMemo(() => ['ALL', ...cauTrucTheoBacSi.map(b => b.ten)], [cauTrucTheoBacSi]);
  const danhSachTenPhong = useMemo(() => ['ALL', ...cauTrucTheoPhong.map(p => p.ten)], [cauTrucTheoPhong]);

  const duLieuHienThiBS = useMemo(() => selectedDoctor === 'ALL' ? cauTrucTheoBacSi : cauTrucTheoBacSi.filter(b => b.ten === selectedDoctor), [cauTrucTheoBacSi, selectedDoctor]);
  const duLieuHienThiPhong = useMemo(() => selectedPhong === 'ALL' ? cauTrucTheoPhong : cauTrucTheoPhong.filter(p => p.ten === selectedPhong), [cauTrucTheoPhong, selectedPhong]);

  const setOptimalColumnWidths = (worksheet) => {
    worksheet.getColumn(1).width = 5;
    worksheet.getColumn(2).width = 8;
    worksheet.getColumn(3).width = 30;
    for (let i = 4; i <= 10; i++) {
      worksheet.getColumn(i).width = 24;
    }
  };

  const applyHeaderStyle = (row, boldFont, thinBorder, headerFill) => {
    row.eachCell((cell, colNumber) => {
      if (colNumber > 1) {
        cell.font = boldFont;
        cell.fill = headerFill;
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: false };
        cell.border = thinBorder;
      }
    });
  };

  const handleExportExcelMauChuan = async () => {
    const workbook = new ExcelJS.Workbook();
    const thinBorder = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
    
    const defaultFont = { name: 'Times New Roman', size: 11 };
    const boldFont = { name: 'Times New Roman', size: 11, bold: true };
    const titleFont = { name: 'Times New Roman', size: 13, bold: true };
    const subTitleFont = { name: 'Times New Roman', size: 11, italic: true };
    const headerFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'D9D9D9' } };

    if (cheDoXem === 'BAC_SI') {
      if (cauTrucTheoBacSi.length === 0) { alert("⚠️ Chưa có dữ liệu Bác sĩ!"); return; }
      
      const wsTongHop = workbook.addWorksheet('Tổng Hợp Bác Sĩ');
      wsTongHop.addRow(["", "SỞ Y TẾ CÀ MAU"]).font = defaultFont;
      wsTongHop.addRow(["", "PKĐK Y ĐỨC SÀI GÒN"]).font = defaultFont;
      wsTongHop.addRow([]);
      
      wsTongHop.mergeCells('B4:H4');
      wsTongHop.getCell('B4').value = 'BÁO CÁO TỔNG HỢP DOANH THU THEO BÁC SĨ';
      wsTongHop.getCell('B4').font = titleFont;
      wsTongHop.getCell('B4').alignment = { horizontal: 'center', vertical: 'middle' };

      wsTongHop.mergeCells('B5:H5');
      wsTongHop.getCell('B5').value = ngayThangBaoCao;
      wsTongHop.getCell('B5').font = subTitleFont;
      wsTongHop.getCell('B5').alignment = { horizontal: 'center', vertical: 'middle' };
      wsTongHop.addRow([]);
      
      const hRow = wsTongHop.addRow(["", "STT", "TÊN BÁC SĨ", "CÔNG KHÁM", "THUỐC & VTYT", "THỦ THUẬT", "XÉT NGHIỆM / CDHA", "TỔNG CỘNG"]);
      applyHeaderStyle(hRow, boldFont, thinBorder, headerFill);

      duLieuHienThiBS.forEach((b, idx) => {
        let tongCongBS = 0;
        let cKham = 0, thuoc = 0, thuthuat = 0, xnCdha = 0;
        ['Công khám', 'Thuốc và VTYT', 'Thủ thuật - Phẫu thuật - VLTL'].forEach(nk => {
          const list = b.cacNhóm[nk] || [];
          const sum = list.reduce((s, x) => s + x.thanhTien, 0);
          if (nk === 'Công khám') cKham += sum;
          if (nk === 'Thuốc và VTYT') thuoc += sum;
          if (nk === 'Thủ thuật - Phẫu thuật - VLTL') thuthuat += sum;
          tongCongBS += sum;
        });
        ['CDHA', 'Xét nghiệm'].forEach(nk => {
          const list = b.cacNhóm[nk] || [];
          const sum = list.reduce((s, x) => s + x.thanhTien, 0);
          xnCdha += sum;
          tongCongBS += sum;
        });
        const r = wsTongHop.addRow(["", idx + 1, b.ten, cKham, thuoc, thuthuat, xnCdha, tongCongBS]);
        r.eachCell((c, colNumber) => {
          c.font = defaultFont;
          c.border = thinBorder;
          if (colNumber >= 4) c.numFmt = '#,##0';
          if (colNumber === 2) c.alignment = { horizontal: 'center', vertical: 'middle' };
          if (colNumber === 3) c.alignment = { horizontal: 'left', vertical: 'middle' };
        });
      });
      setOptimalColumnWidths(wsTongHop);

      duLieuHienThiBS.forEach(b => {
        const sheetName = String(b.ten).substring(0, 30).replace(/[:\\\/?*\[\]]/g, '');
        const wsBS = workbook.addWorksheet(sheetName);
        wsBS.addRow(["", "SỞ Y TẾ CÀ MAU"]).font = defaultFont;
        wsBS.addRow(["", "PKĐK Y ĐỨC SÀI GÒN"]).font = defaultFont;
        wsBS.addRow([]);
        
        wsBS.mergeCells('B4:G4');
        wsBS.getCell('B4').value = `CHI TIẾT CHỈ ĐỊNH - BÁC SĨ: ${b.ten.toUpperCase()}`;
        wsBS.getCell('B4').font = titleFont;
        wsBS.getCell('B4').alignment = { horizontal: 'center', vertical: 'middle' };

        wsBS.mergeCells('B5:G5');
        wsBS.getCell('B5').value = ngayThangBaoCao;
        wsBS.getCell('B5').font = subTitleFont;
        wsBS.getCell('B5').alignment = { horizontal: 'center', vertical: 'middle' };
        wsBS.addRow([]);

        const sharedHeader = wsBS.addRow(["", "STT", "TÊN DỊCH VỤ", "LOẠI", "SỐ LƯỢNG", "ĐƠN GIÁ", "THÀNH TIỀN"]);
        applyHeaderStyle(sharedHeader, boldFont, thinBorder, headerFill);

        ['Công khám', 'CDHA', 'Xét nghiệm', 'Thuốc và VTYT', 'Thủ thuật - Phẫu thuật - VLTL'].forEach(nhomKey => {
          const list = b.cacNhóm[nhomKey] || [];
          if (list.length === 0) return;
          
          const groupRow = wsBS.addRow(["", "", nhomKey.toUpperCase()]);
          groupRow.font = boldFont;
          groupRow.eachCell((c) => { c.border = thinBorder; });
          groupRow.getCell(3).alignment = { horizontal: 'left', vertical: 'middle' };

          let tongTienNhom = 0;
          list.forEach((dv, i) => {
            tongTienNhom += dv.thanhTien;
            const dvRow = wsBS.addRow(["", i + 1, dv.tenDichVu, dv.loai, dv.soLuong, dv.donGia, dv.thanhTien]);
            dvRow.eachCell((c, colNumber) => {
              c.font = defaultFont;
              c.border = thinBorder;
              if (colNumber === 2 || colNumber === 4 || colNumber === 5) c.alignment = { horizontal: 'center', vertical: 'middle' };
              if (colNumber === 3) c.alignment = { horizontal: 'left', vertical: 'middle' };
              if (colNumber === 6 || colNumber === 7) c.numFmt = '#,##0';
            });
          });

          const totalRow = wsBS.addRow(["", "", "Tổng cộng", "", "", "", tongTienNhom]);
          totalRow.eachCell((c, colNumber) => {
            c.font = boldFont;
            c.border = thinBorder;
            if (colNumber === 3) c.alignment = { horizontal: 'left', vertical: 'middle' };
            if (colNumber === 7) {
              c.alignment = { horizontal: 'right', vertical: 'middle' };
              c.numFmt = '#,##0';
            }
          });
        });
        setOptimalColumnWidths(wsBS);
      });

    } else {
      if (cauTrucTheoPhong.length === 0) { alert("⚠️ Chưa có dữ liệu Phòng ban!"); return; }
      
      const wsTongHop = workbook.addWorksheet('Tổng Hợp Phòng Ban');
      wsTongHop.addRow(["", "SỞ Y TẾ CÀ MAU"]).font = defaultFont;
      wsTongHop.addRow(["", "PKĐK Y ĐỨC SÀI GÒN"]).font = defaultFont;
      wsTongHop.addRow([]);
      
      wsTongHop.mergeCells('B4:H4');
      wsTongHop.getCell('B4').value = 'BÁO CÁO TỔNG HỢP DOANH THU THEO PHÒNG';
      wsTongHop.getCell('B4').font = titleFont;
      wsTongHop.getCell('B4').alignment = { horizontal: 'center', vertical: 'middle' };

      wsTongHop.mergeCells('B5:H5');
      wsTongHop.getCell('B5').value = ngayThangBaoCao;
      wsTongHop.getCell('B5').font = subTitleFont;
      wsTongHop.getCell('B5').alignment = { horizontal: 'center', vertical: 'middle' };
      wsTongHop.addRow([]);
      
      const hRow = wsTongHop.addRow(["", "STT", "TÊN PHÒNG BAN", "CÔNG KHÁM", "THUỐC & VTYT", "THỦ THUẬT", "XÉT NGHIỆM / CDHA", "TỔNG CỘNG"]);
      applyHeaderStyle(hRow, boldFont, thinBorder, headerFill);

      duLieuHienThiPhong.forEach((p, idx) => {
        let tongCongPhong = 0;
        let cKham = 0, thuoc = 0, thuthuat = 0, xnCdha = 0;
        ['Công khám', 'Thuốc và VTYT', 'Thủ thuật - Phẫu thuật - VLTL'].forEach(nk => {
          const list = p.cacNhóm[nk] || [];
          const sum = list.reduce((s, x) => s + x.thanhTien, 0);
          if (nk === 'Công khám') cKham += sum;
          if (nk === 'Thuốc và VTYT') thuoc += sum;
          if (nk === 'Thủ thuật - Phẫu thuật - VLTL') thuthuat += sum;
          tongCongPhong += sum;
        });
        ['CDHA', 'Xét nghiệm'].forEach(nk => {
          const list = p.cacNhóm[nk] || [];
          const sum = list.reduce((s, x) => s + x.thanhTien, 0);
          xnCdha += sum;
          tongCongPhong += sum;
        });
        const r = wsTongHop.addRow(["", idx + 1, p.ten, cKham, thuoc, thuthuat, xnCdha, tongCongPhong]);
        r.eachCell((c, colNumber) => {
          c.font = defaultFont;
          c.border = thinBorder;
          if (colNumber >= 4) c.numFmt = '#,##0';
          if (colNumber === 2) c.alignment = { horizontal: 'center', vertical: 'middle' };
          if (colNumber === 3) c.alignment = { horizontal: 'left', vertical: 'middle' };
        });
      });
      setOptimalColumnWidths(wsTongHop);

      duLieuHienThiPhong.forEach(p => {
        const sheetName = String(p.ten).substring(0, 30).replace(/[:\\\/?*\[\]]/g, '');
        const wsPhong = workbook.addWorksheet(sheetName);
        wsPhong.addRow(["", "SỞ Y TẾ CÀ MAU"]).font = defaultFont;
        wsPhong.addRow(["", "PKĐK Y ĐỨC SÀI GÒN"]).font = defaultFont;
        wsPhong.addRow([]);
        
        wsPhong.mergeCells('B4:G4');
        wsPhong.getCell('B4').value = `CHI TIẾT CHỈ ĐỊNH - PHÒNG BAN: ${p.ten.toUpperCase()}`;
        wsPhong.getCell('B4').font = titleFont;
        wsPhong.getCell('B4').alignment = { horizontal: 'center', vertical: 'middle' };

        wsPhong.mergeCells('B5:G5');
        wsPhong.getCell('B5').value = ngayThangBaoCao;
        wsPhong.getCell('B5').font = subTitleFont;
        wsPhong.getCell('B5').alignment = { horizontal: 'center', vertical: 'middle' };
        wsPhong.addRow([]);

        const sharedHeader = wsPhong.addRow(["", "STT", "TÊN DỊCH VỤ", "LOẠI", "SỐ LƯỢNG", "ĐƠN GIÁ", "THÀNH TIỀN"]);
        applyHeaderStyle(sharedHeader, boldFont, thinBorder, headerFill);

        ['Công khám', 'CDHA', 'Xét nghiệm', 'Thuốc và VTYT', 'Thủ thuật - Phẫu thuật - VLTL'].forEach(nhomKey => {
          const list = p.cacNhóm[nhomKey] || [];
          if (list.length === 0) return;
          
          const groupRow = wsPhong.addRow(["", "", nhomKey.toUpperCase()]);
          groupRow.font = boldFont;
          groupRow.eachCell((c) => { c.border = thinBorder; });
          groupRow.getCell(3).alignment = { horizontal: 'left', vertical: 'middle' };

          let tongTienNhom = 0;
          list.forEach((dv, i) => {
            tongTienNhom += dv.thanhTien;
            const dvRow = wsPhong.addRow(["", i + 1, dv.tenDichVu, dv.loai, dv.soLuong, dv.donGia, dv.thanhTien]);
            dvRow.eachCell((c, colNumber) => {
              c.font = defaultFont;
              c.border = thinBorder;
              if (colNumber === 2 || colNumber === 4 || colNumber === 5) c.alignment = { horizontal: 'center', vertical: 'middle' };
              if (colNumber === 3) c.alignment = { horizontal: 'left', vertical: 'middle' };
              if (colNumber === 6 || colNumber === 7) c.numFmt = '#,##0';
            });
          });

          const totalRow = wsPhong.addRow(["", "", "Tổng cộng", "", "", "", tongTienNhom]);
          totalRow.eachCell((c, colNumber) => {
            c.font = boldFont;
            c.border = thinBorder;
            if (colNumber === 3) c.alignment = { horizontal: 'left', vertical: 'middle' };
            if (colNumber === 7) {
              c.alignment = { horizontal: 'right', vertical: 'middle' };
              c.numFmt = '#,##0';
            }
          });
        });
        setOptimalColumnWidths(wsPhong);
      });
    }

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = window.URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = cheDoXem === 'BAC_SI' ? `TONG_HOP_CLS_THUOC_BHYT_BS.xlsx` : `TONG_HOP_CLS_THUOC_BHYT_PHONG.xlsx`;
    document.body.appendChild(anchor); anchor.click(); document.body.removeChild(anchor);
    window.URL.revokeObjectURL(url);
  };

  return (
    <div style={{ background: '#f8fafc', padding: '24px', borderRadius: '12px', border: '1px solid #e2e8f0', fontFamily: 'Segoe UI, Roboto, sans-serif' }}>
      
      {/* HEADER TITLE */}
      <div style={{ textAlign: 'center', marginBottom: '20px', paddingBottom: '12px', borderBottom: '2px solid #e2e8f0', backgroundColor: 'white', padding: '16px', borderRadius: '8px', boxShadow: '0 2px 4px rgba(0,0,0,0.02)' }}>
        <h3 style={{ margin: '0 0 5px 0', fontSize: '18px', fontWeight: '800', color: '#1e293b', textTransform: 'uppercase', letterSpacing: '0.5px' }}>
          {cheDoXem === 'BAC_SI' ? '📊 TỔNG HỢP CHỈ ĐỊNH & THUỐC THEO BÁC SĨ' : '🏥 TỔNG HỢP CHỈ ĐỊNH & THUỐC THEO PHÒNG BAN'}
        </h3>
        <span style={{ fontSize: '13px', fontWeight: '700', color: '#4f46e5', backgroundColor: '#e0e7ff', padding: '4px 12px', borderRadius: '20px' }}>
          {ngayThangBaoCao}
        </span>
      </div>

      {/* TOGGLE BUTTONS */}
      <div style={{ display: 'flex', justifyContent: 'center', gap: '15px', marginBottom: '20px' }}>
        <button onClick={() => setCheDoXem('BAC_SI')} style={{ padding: '10px 24px', backgroundColor: cheDoXem === 'BAC_SI' ? '#4f46e5' : '#ffffff', color: cheDoXem === 'BAC_SI' ? 'white' : '#475569', border: cheDoXem === 'BAC_SI' ? 'none' : '1px solid #cbd5e1', borderRadius: '8px', fontWeight: 'bold', cursor: 'pointer' }}>
          👨‍⚕️ Xem theo Bác sĩ
        </button>
        <button onClick={() => setCheDoXem('PHONG')} style={{ padding: '10px 24px', backgroundColor: cheDoXem === 'PHONG' ? '#0284c7' : '#ffffff', color: cheDoXem === 'PHONG' ? 'white' : '#475569', border: cheDoXem === 'PHONG' ? 'none' : '1px solid #cbd5e1', borderRadius: '8px', fontWeight: 'bold', cursor: 'pointer' }}>
          🏥 Xem theo Phòng ban
        </button>
      </div>

      {/* FILTER & EXPORT BAR */}
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '24px', backgroundColor: 'white', padding: '16px 20px', borderRadius: '8px', border: '1px solid #e2e8f0', flexWrap: 'wrap', gap: '15px' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
          <span style={{ fontSize: '14px', fontWeight: '700', color: '#334155' }}>
            {cheDoXem === 'BAC_SI' ? '🔍 Lọc Bác sĩ:' : '🔍 Lọc Phòng ban:'}
          </span>
          {cheDoXem === 'BAC_SI' ? (
            <select value={selectedDoctor} onChange={(e) => setSelectedDoctor(e.target.value)} style={{ padding: '9px 14px', borderRadius: '8px', border: '1px solid #cbd5e1', fontSize: '14px', minWidth: '260px', fontWeight: '700', color: '#4f46e5', outline: 'none', backgroundColor: '#f8fafc' }}>
              {danhSachTenBS.map((d, i) => <option key={i} value={d}>{d === 'ALL' ? '--- Tất cả Bác sĩ ---' : d}</option>)}
            </select>
          ) : (
            <select value={selectedPhong} onChange={(e) => setSelectedPhong(e.target.value)} style={{ padding: '9px 14px', borderRadius: '8px', border: '1px solid #cbd5e1', fontSize: '14px', minWidth: '260px', fontWeight: '700', color: '#0284c7', outline: 'none', backgroundColor: '#f8fafc' }}>
              {danhSachTenPhong.map((p, i) => <option key={i} value={p}>{p === 'ALL' ? '--- Tất cả Phòng ban ---' : p}</option>)}
            </select>
          )}
        </div>

        <button onClick={handleExportExcelMauChuan} style={{ padding: '10px 20px', backgroundColor: '#10b981', color: 'white', border: 'none', borderRadius: '8px', fontWeight: 'bold', cursor: 'pointer', display: 'flex', alignItems: 'center', gap: '8px' }}>
          <span>📥</span> Xuất Excel Báo Cáo Chuẩn
        </button>
      </div>

      {/* CONTENT LIST */}
      <div style={{ maxHeight: '600px', overflowY: 'auto', display: 'flex', flexDirection: 'column', gap: '20px', paddingRight: '4px' }}>
        {cheDoXem === 'BAC_SI' ? (
          duLieuHienThiBS.length === 0 ? (
            <div style={{ textAlign: 'center', padding: '50px', color: '#64748b', backgroundColor: 'white', borderRadius: '8px', border: '1px solid #e2e8f0' }}>❌ Chưa có dữ liệu Bác sĩ.</div>
          ) : (
            duLieuHienThiBS.map((b, idx) => (
              <div key={idx} style={{ background: 'white', padding: '20px', borderRadius: '10px', border: '1px solid #cbd5e1', boxShadow: '0 4px 12px rgba(0,0,0,0.03)' }}>
                <div style={{ fontSize: '16px', fontWeight: '800', color: '#1e293b', borderBottom: '2px solid #4f46e5', paddingBottom: '8px', marginBottom: '15px' }}>
                  🧑‍⚕️ Bác sĩ: <span style={{ color: '#4f46e5' }}>{b.ten}</span>
                </div>
                {['Công khám', 'CDHA', 'Xét nghiệm', 'Thuốc và VTYT', 'Thủ thuật - Phẫu thuật - VLTL'].map(nhomKey => {
                  const list = b.cacNhóm[nhomKey] || [];
                  if (list.length === 0) return null;
                  const tongTienNhom = list.reduce((s, d) => s + d.thanhTien, 0);

                  return (
                    <div key={nhomKey} style={{ marginBottom: '15px', backgroundColor: '#f8fafc', borderRadius: '8px', border: '1px solid #e2e8f0', overflow: 'hidden' }}>
                      <div style={{ fontSize: '13px', fontWeight: '800', color: '#334155', backgroundColor: '#f1f5f9', padding: '10px 14px', borderBottom: '1px solid #e2e8f0', display: 'flex', justifyContent: 'space-between' }}>
                        <span>📁 {nhomKey} ({list.length} mục)</span>
                        <span style={{ color: '#059669' }}>Tổng: {tongTienNhom.toLocaleString('vi-VN')} đ</span>
                      </div>
                      <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12.5px' }}>
                        <thead>
                          <tr style={{ backgroundColor: '#ffffff', color: '#64748b', borderBottom: '1px solid #e2e8f0', textAlign: 'left' }}>
                            <th style={{ padding: '8px 12px' }}>TÊN DỊCH VỤ / THUỐC</th>
                            <th style={{ padding: '8px 12px', width: '70px', textAlign: 'center' }}>LOẠI</th>
                            <th style={{ padding: '8px 12px', width: '70px', textAlign: 'center' }}>SL</th>
                            <th style={{ padding: '8px 12px', width: '110px', textAlign: 'right' }}>ĐƠN GIÁ</th>
                            <th style={{ padding: '8px 12px', width: '130px', textAlign: 'right' }}>THÀNH TIỀN</th>
                          </tr>
                        </thead>
                        <tbody>
                          {list.map((dv, i) => (
                            <tr key={i} style={{ borderBottom: '1px solid #f1f5f9', backgroundColor: i % 2 === 0 ? '#ffffff' : '#fafafa' }}>
                              <td style={{ padding: '8px 12px', fontWeight: '600', color: '#1e293b' }}>{dv.tenDichVu}</td>
                              <td style={{ padding: '8px 12px', textAlign: 'center', color: '#475569' }}>{dv.loai}</td>
                              <td style={{ padding: '8px 12px', textAlign: 'center', fontWeight: 'bold', color: '#2563eb' }}>{dv.soLuong}</td>
                              <td style={{ padding: '8px 12px', textAlign: 'right', color: '#475569' }}>{dv.donGia.toLocaleString('vi-VN')} đ</td>
                              <td style={{ padding: '8px 12px', textAlign: 'right', fontWeight: 'bold', color: '#059669' }}>{dv.thanhTien.toLocaleString('vi-VN')} đ</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  );
                })}
              </div>
            ))
          )
        ) : (
          duLieuHienThiPhong.length === 0 ? (
            <div style={{ textAlign: 'center', padding: '50px', color: '#64748b', backgroundColor: 'white', borderRadius: '8px', border: '1px solid #e2e8f0' }}>❌ Chưa có dữ liệu Phòng ban.</div>
          ) : (
            duLieuHienThiPhong.map((p, idx) => (
              <div key={idx} style={{ background: 'white', padding: '20px', borderRadius: '10px', border: '1px solid #cbd5e1', boxShadow: '0 4px 12px rgba(0,0,0,0.03)' }}>
                <div style={{ fontSize: '16px', fontWeight: '800', color: '#1e293b', borderBottom: '2px solid #0284c7', paddingBottom: '8px', marginBottom: '15px' }}>
                  🏥 Phòng ban: <span style={{ color: '#0284c7' }}>{p.ten}</span>
                </div>
                {['Công khám', 'CDHA', 'Xét nghiệm', 'Thuốc và VTYT', 'Thủ thuật - Phẫu thuật - VLTL'].map(nhomKey => {
                  const list = p.cacNhóm[nhomKey] || [];
                  if (list.length === 0) return null;
                  const tongTienNhom = list.reduce((s, d) => s + d.thanhTien, 0);

                  return (
                    <div key={nhomKey} style={{ marginBottom: '15px', backgroundColor: '#f8fafc', borderRadius: '8px', border: '1px solid #e2e8f0', overflow: 'hidden' }}>
                      <div style={{ fontSize: '13px', fontWeight: '800', color: '#334155', backgroundColor: '#f1f5f9', padding: '10px 14px', borderBottom: '1px solid #e2e8f0', display: 'flex', justifyContent: 'space-between' }}>
                        <span>📁 {nhomKey} ({list.length} mục)</span>
                        <span style={{ color: '#059669' }}>Tổng: {tongTienNhom.toLocaleString('vi-VN')} đ</span>
                      </div>
                      <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12.5px' }}>
                        <thead>
                          <tr style={{ backgroundColor: '#ffffff', color: '#64748b', borderBottom: '1px solid #e2e8f0', textAlign: 'left' }}>
                            <th style={{ padding: '8px 12px' }}>TÊN DỊCH VỤ / THUỐC</th>
                            <th style={{ padding: '8px 12px', width: '70px', textAlign: 'center' }}>LOẠI</th>
                            <th style={{ padding: '8px 12px', width: '70px', textAlign: 'center' }}>SL</th>
                            <th style={{ padding: '8px 12px', width: '110px', textAlign: 'right' }}>ĐƠN GIÁ</th>
                            <th style={{ padding: '8px 12px', width: '130px', textAlign: 'right' }}>THÀNH TIỀN</th>
                          </tr>
                        </thead>
                        <tbody>
                          {list.map((dv, i) => (
                            <tr key={i} style={{ borderBottom: '1px solid #f1f5f9', backgroundColor: i % 2 === 0 ? '#ffffff' : '#fafafa' }}>
                              <td style={{ padding: '8px 12px', fontWeight: '600', color: '#1e293b' }}>{dv.tenDichVu}</td>
                              <td style={{ padding: '8px 12px', textAlign: 'center', color: '#475569' }}>{dv.loai}</td>
                              <td style={{ padding: '8px 12px', textAlign: 'center', fontWeight: 'bold', color: '#2563eb' }}>{dv.soLuong}</td>
                              <td style={{ padding: '8px 12px', textAlign: 'right', color: '#475569' }}>{dv.donGia.toLocaleString('vi-VN')} đ</td>
                              <td style={{ padding: '8px 12px', textAlign: 'right', fontWeight: 'bold', color: '#059669' }}>{dv.thanhTien.toLocaleString('vi-VN')} đ</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  );
                })}
              </div>
            ))
          )
        )}
      </div>

    </div>
  );
}

export default TongHopYLenhBacSi;
import React, { useState } from 'react';
import LocHoaDonCtyMST from './LocHoaDonCtyMST';
import LocHoaDonCongTy from './LocHoaDonCongTy';
import TongHopYLenhBacSi from './TongHopYLenhBacSi';
import LocTrungThuChi from './LocTrungThuChi';
import TonKhoThuoc from './TonKhoThuoc';

export default function FinanceApp() {
  const [activeTab, setActiveTab] = useState('hoadon-mst');

  return (
    <div style={{ fontFamily: 'Segoe UI, sans-serif', background: '#f1f5f9', minHeight: '100vh', padding: '20px' }}>
      
      {/* THANH HEADER ĐIỀU HƯỚNG */}
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', background: 'white', padding: '16px 24px', borderRadius: '10px', boxShadow: '0 2px 4px rgba(0,0,0,0.05)', marginBottom: '20px' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
          <span style={{ backgroundColor: '#db2777', color: 'white', padding: '4px 10px', borderRadius: '6px', fontWeight: 'bold' }}>💰</span>
          <h2 style={{ margin: 0, color: '#1e293b', fontSize: '18px' }}>medOS - Phân Hệ Quản Lý Tài Chính & Dược</h2>
        </div>
        
        {/* Nút liên kết quay lại app Chấm công chính */}
        <a 
          href="https://myxu.vercel.app" 
          target="_blank" 
          rel="noopener noreferrer" 
          style={{ textDecoration: 'none', padding: '8px 16px', background: '#334155', color: 'white', borderRadius: '6px', fontWeight: 'bold', fontSize: '12.5px' }}
        >
          ← Quay lại App Chấm Công medOS
        </a>
      </div>

      {/* THANH MENU TAB CHỌN CHỨC NĂNG */}
      <div style={{ display: 'flex', gap: '10px', marginBottom: '20px', flexWrap: 'wrap' }}>
        <button 
          onClick={() => setActiveTab('hoadon-mst')}
          style={{ padding: '9px 16px', background: activeTab === 'hoadon-mst' ? '#2563eb' : 'white', color: activeTab === 'hoadon-mst' ? 'white' : '#334155', border: '1px solid #cbd5e1', borderRadius: '6px', fontWeight: 'bold', cursor: 'pointer', fontSize: '13px' }}
        >
          🏢 Lọc Hóa Đơn Cty & MST
        </button>
        <button 
          onClick={() => setActiveTab('hoadon-cty')}
          style={{ padding: '9px 16px', background: activeTab === 'hoadon-cty' ? '#2563eb' : 'white', color: activeTab === 'hoadon-cty' ? 'white' : '#334155', border: '1px solid #cbd5e1', borderRadius: '6px', fontWeight: 'bold', cursor: 'pointer', fontSize: '13px' }}
        >
          📊 Phân Tích Hóa Đơn Công Ty
        </button>
        <button 
          onClick={() => setActiveTab('tonghop-bs')}
          style={{ padding: '9px 16px', background: activeTab === 'tonghop-bs' ? '#2563eb' : 'white', color: activeTab === 'tonghop-bs' ? 'white' : '#334155', border: '1px solid #cbd5e1', borderRadius: '6px', fontWeight: 'bold', cursor: 'pointer', fontSize: '13px' }}
        >
          📑 Tổng Hợp Y Lệnh / Doanh Thu
        </button>
        <button 
          onClick={() => setActiveTab('trung-thuchi')}
          style={{ padding: '9px 16px', background: activeTab === 'trung-thuchi' ? '#2563eb' : 'white', color: activeTab === 'trung-thuchi' ? 'white' : '#334155', border: '1px solid #cbd5e1', borderRadius: '6px', fontWeight: 'bold', cursor: 'pointer', fontSize: '13px' }}
        >
          🔍 Lọc Trùng Thu Chi
        </button>
        <button 
          onClick={() => setActiveTab('tonkho')}
          style={{ padding: '9px 16px', background: activeTab === 'tonkho' ? '#2563eb' : 'white', color: activeTab === 'tonkho' ? 'white' : '#334155', border: '1px solid #cbd5e1', borderRadius: '6px', fontWeight: 'bold', cursor: 'pointer', fontSize: '13px' }}
        >
          📦 Quản Lý Tồn Kho Thuốc
        </button>
      </div>

      {/* KHU VỰC HIỂN THỊ NỘI DUNG TÍNH NĂNG */}
      <div style={{ background: 'white', padding: '24px', borderRadius: '10px', boxShadow: '0 4px 6px rgba(0,0,0,0.02)', minHeight: '600px' }}>
        {activeTab === 'hoadon-mst' && <LocHoaDonCtyMST />}
        {activeTab === 'hoadon-cty' && <LocHoaDonCongTy />}
        {activeTab === 'tonghop-bs' && <TongHopYLenhBacSi />}
        {activeTab === 'trung-thuchi' && <LocTrungThuChi />}
        {activeTab === 'tonkho' && <TonKhoThuoc />}
      </div>

    </div>
  );
}
"use client";

import { useState } from 'react';

export default function HomePage() {
  const [downloading, setDownloading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const handleDownload = async () => {
    try {
      setDownloading(true);
      setError(null);
      const res = await fetch('/api/export', { method: 'GET' });
      if (!res.ok) throw new Error('Failed to generate workbook');
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'Logistics_Dashboard.xlsx';
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
    } catch (e: any) {
      setError(e?.message ?? 'Unexpected error');
    } finally {
      setDownloading(false);
    }
  };

  return (
    <main style={{
      minHeight: '100vh',
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'center',
      background: 'linear-gradient(135deg, #0ea5e9 0%, #22d3ee 100%)',
      padding: 24
    }}>
      <div style={{
        width: '100%',
        maxWidth: 720,
        background: 'white',
        borderRadius: 16,
        boxShadow: '0 20px 60px rgba(0,0,0,0.15)',
        padding: 32
      }}>
        <h1 style={{ fontSize: 28, marginBottom: 8 }}>Logistics Operations Excel Dashboard</h1>
        <p style={{ color: '#475569', marginBottom: 24 }}>
          Download a pre-built Excel workbook with an interactive dashboard. Choose month and vendor inside Excel to view KPIs: On-time vs Delayed, Delay reasons, Truck types, Trucks used, and breakdown counts.
        </p>
        <button onClick={handleDownload} disabled={downloading} style={{
          background: '#0ea5e9',
          color: 'white',
          border: 0,
          borderRadius: 12,
          padding: '12px 18px',
          fontSize: 16,
          cursor: 'pointer'
        }}>
          {downloading ? 'Generating?' : 'Download Excel Dashboard'}
        </button>
        {error && (
          <p style={{ color: '#ef4444', marginTop: 16 }}>{error}</p>
        )}
        <div style={{ marginTop: 24, fontSize: 12, color: '#64748b' }}>
          Tip: Open in modern Excel (Microsoft 365) for best dynamic formulas support.
        </div>
      </div>
    </main>
  );
}

'use client';

import { useState } from 'react';
import './globals.css';

export default function ExcelFillerPage() {
    const [formData, setFormData] = useState({
        branches: '1',
        product: 'Premium Health Checkup',
        units: '100',
        price: '2500',
    });
    const [loading, setLoading] = useState(false);
    const [status, setStatus] = useState('');

    const handleChange = (e) => {
        setFormData({ ...formData, [e.target.name]: e.target.value });
    };

    const handleGenerate = async () => {
        setLoading(true);
        setStatus('Generating your Excel file...');

        try {
            const response = await fetch('/api/fill-excel', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(formData),
            });

            if (!response.ok) throw new Error('Generation failed');

            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `Plan_${formData.product.replace(/\s+/g, '_')}.xlsx`;
            document.body.appendChild(a);
            a.click();
            a.remove();

            setStatus('✅ Success! File downloaded.');
        } catch (error) {
            console.error(error);
            setStatus('❌ Error: ' + error.message);
        } finally {
            setLoading(false);
        }
    };

    return (
        <div className="card">
            <h1>📊 Business Plan Filler</h1>

            <div className="form-group">
                <label>Number of Branches (H7)</label>
                <input
                    type="number"
                    name="branches"
                    value={formData.branches}
                    onChange={handleChange}
                />
            </div>

            <div className="form-group">
                <label>Product Name (F10)</label>
                <input
                    type="text"
                    name="product"
                    value={formData.product}
                    onChange={handleChange}
                />
            </div>

            <div className="form-group">
                <label>Monthly Units per Branch (H10)</label>
                <input
                    type="number"
                    name="units"
                    value={formData.units}
                    onChange={handleChange}
                />
            </div>

            <div className="form-group">
                <label>Sale Price (J10)</label>
                <input
                    type="number"
                    name="price"
                    value={formData.price}
                    onChange={handleChange}
                />
            </div>

            <button onClick={handleGenerate} disabled={loading}>
                {loading ? 'Processing...' : 'Generate & Download XLSX'}
            </button>

            {status && <div className="status">{status}</div>}
        </div>
    );
}

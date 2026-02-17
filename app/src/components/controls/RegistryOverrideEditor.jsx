/**
 * RegistryOverrideEditor — Dynamic table for advanced.registryOverrides.
 */
export default function RegistryOverrideEditor({ value = [], onChange }) {
  const overrides = [...value];

  const handleAdd = () => {
    onChange([...overrides, { path: '', name: '', value: '', type: 'DWord' }]);
  };

  const handleRemove = (index) => {
    onChange(overrides.filter((_, i) => i !== index));
  };

  const handleFieldChange = (index, field, val) => {
    const next = [...overrides];
    next[index] = { ...next[index], [field]: val };
    onChange(next);
  };

  const inputStyle = {
    background: 'var(--bg-tertiary)', border: '1px solid var(--border-color)',
    borderRadius: 'var(--radius-sm)', padding: '4px 8px', color: 'var(--text-primary)',
    fontSize: 12, fontFamily: 'var(--font-mono)'
  };

  return (
    <div>
      <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
        <thead>
          <tr style={{ color: 'var(--text-muted)', textAlign: 'left' }}>
            <th style={{ padding: '4px 6px' }}>Registry Path</th>
            <th style={{ padding: '4px 6px' }}>Name</th>
            <th style={{ padding: '4px 6px', width: 100 }}>Value</th>
            <th style={{ padding: '4px 6px', width: 90 }}>Type</th>
            <th style={{ width: 30 }} />
          </tr>
        </thead>
        <tbody>
          {overrides.map((o, i) => (
            <tr key={i} style={{ borderTop: '1px solid var(--border-color)' }}>
              <td style={{ padding: 4 }}>
                <input style={{ ...inputStyle, width: '100%' }} value={o.path}
                  onChange={(e) => handleFieldChange(i, 'path', e.target.value)}
                  placeholder="HKCU:\..." />
              </td>
              <td style={{ padding: 4 }}>
                <input style={{ ...inputStyle, width: '100%' }} value={o.name}
                  onChange={(e) => handleFieldChange(i, 'name', e.target.value)} />
              </td>
              <td style={{ padding: 4 }}>
                <input style={{ ...inputStyle, width: '100%' }} value={o.value}
                  onChange={(e) => handleFieldChange(i, 'value', e.target.value)} />
              </td>
              <td style={{ padding: 4 }}>
                <select style={{ ...inputStyle, width: '100%' }} value={o.type}
                  onChange={(e) => handleFieldChange(i, 'type', e.target.value)}>
                  <option value="DWord">DWord</option>
                  <option value="String">String</option>
                  <option value="ExpandString">ExpandString</option>
                  <option value="Binary">Binary</option>
                </select>
              </td>
              <td style={{ padding: 4, textAlign: 'center' }}>
                <button className="btn btn--danger" onClick={() => handleRemove(i)}
                  style={{ padding: '2px 6px', fontSize: 11 }}>✕</button>
              </td>
            </tr>
          ))}
        </tbody>
      </table>
      <button className="btn" onClick={handleAdd} style={{ marginTop: 8, fontSize: 11 }}>
        + Add Override
      </button>
    </div>
  );
}

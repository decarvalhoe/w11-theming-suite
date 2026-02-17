/**
 * FilePicker — Opens a native file dialog and displays the selected path.
 * @param {string} value       – Current file path
 * @param {function} onChange   – Callback with the new path
 * @param {array} filters       – Electron dialog filters, e.g. [{ name: 'Images', extensions: ['png','jpg'] }]
 * @param {string} placeholder  – Placeholder text when no file is selected
 */
export default function FilePicker({ value, onChange, filters, placeholder = 'No file selected' }) {
  const handleClick = async () => {
    if (!window.themeAPI?.openFile) return;
    const filePath = await window.themeAPI.openFile(filters);
    if (filePath) {
      onChange(filePath);
    }
  };

  const fileName = value ? value.split(/[/\\]/).pop() : '';

  return (
    <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
      <button type="button" className="btn" onClick={handleClick} style={{ fontSize: 11, padding: '4px 10px' }}>
        Browse...
      </button>
      <span style={{
        fontSize: 12, color: value ? 'var(--text-secondary)' : 'var(--text-muted)',
        maxWidth: 180, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap'
      }} title={value || ''}>
        {fileName || placeholder}
      </span>
    </div>
  );
}

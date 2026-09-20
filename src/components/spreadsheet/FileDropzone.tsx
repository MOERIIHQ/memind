import { useState, type DragEvent, type ChangeEvent } from 'react';
import Button from '../Button';
import Icon from '../Icon';
import './FileDropzone.css';

interface FileDropzoneProps {
  onFileSelect: (file: File) => void;
  onDragDrop?: (file: File) => void;
  accept?: string;
  maxSizeMB?: number;
}

const FileDropzone = ({
  onFileSelect,
  onDragDrop,
  accept = '.xlsx,.xls,.csv',
  maxSizeMB = 10,
}: FileDropzoneProps) => {
  const [isDragging, setIsDragging] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const fileInputId = 'dropzone-file-input';

  const validateFile = (file: File): string | null => {
    if (file.size > maxSizeMB * 1024 * 1024) {
      return `File is too large. Maximum size is ${maxSizeMB} MB.`;
    }
    return null;
  };

  const handleFile = (file: File) => {
    const validationError = validateFile(file);
    if (validationError) {
      setError(validationError);
      return;
    }
    setError(null);
    onFileSelect(file);
  };

  const handleDragOver = (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(true);
  };

  const handleDragLeave = (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.currentTarget === e.relatedTarget) {
      setIsDragging(false);
    }
  };

  const handleDrop = (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);

    const files = Array.from(e.dataTransfer.files);
    if (files.length === 0) return;

    const file = files[0];
    onDragDrop?.(file);
    handleFile(file);
  };

  const handleFileInputChange = (e: ChangeEvent<HTMLInputElement>) => {
    console.log('[UPLOAD] FileDropzone handleFileInputChange fired, files count:', e.target?.files?.length || 0);
    const files = e.target.files;
    if (!files || files.length === 0) return;

    const file = files[0];
    console.log('[UPLOAD] FileDropzone file:', file?.name, file?.size, file?.type);
    handleFile(file);
    e.target.value = '';
  };

  return (
    <div
      className={`dropzone ${isDragging ? 'dropzone--dragging' : ''}`}
      onDragOver={handleDragOver}
      onDragLeave={handleDragLeave}
      onDrop={handleDrop}
      aria-describedby="dropzone-instructions"
    >
      <input
        id={fileInputId}
        type="file"
        accept={accept}
        onChange={handleFileInputChange}
        className="dropzone__input"
        aria-label="Select spreadsheet file"
        tabIndex={-1}
        style={{
          position: 'absolute',
          top: 0,
          left: 0,
          width: '100%',
          height: '100%',
          opacity: 0,
          cursor: 'pointer',
          zIndex: 2,
        }}
      />
      <div className="dropzone__content">
        <div className="dropzone__icon" aria-hidden="true">
          <Icon name="upload" size={48} />
        </div>
        <h3 className="dropzone__title">Upload a spreadsheet</h3>
        <p className="dropzone__description" id="dropzone-instructions">
          Drag and drop a file here, or click to browse
        </p>
        <p className="dropzone__formats">
          Supported formats: .xlsx, .xls, .csv (max {maxSizeMB} MB)
        </p>
        {error && (
          <p className="dropzone__error">
            <Icon name="alertCircle" size={14} />
            {error}
          </p>
        )}
        {/* Use label wrapping native input instead of custom Button */}
        <label htmlFor={fileInputId} className="dropzone__button-label">
          <Button variant="primary" size="lg" leftIcon={<Icon name="folderOpen" size={18} />}>
            Choose File
          </Button>
        </label>
      </div>
      {isDragging && <div className="dropzone__overlay" aria-label="Drop to upload"></div>}
    </div>
  );
};

export default FileDropzone;

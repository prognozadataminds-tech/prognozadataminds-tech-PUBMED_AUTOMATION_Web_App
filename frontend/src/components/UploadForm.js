import React, { useState } from 'react';
import { ToastContainer, toast } from 'react-toastify';
import 'react-toastify/dist/ReactToastify.css';
// import axios from 'axios'; // Only needed if using axios version

const UploadForm = () => {
  const [file, setFile] = useState(null);
  const [isLoading, setIsLoading] = useState(false);

  const handleDrop = (e) => {
    e.preventDefault();
    const droppedFile = e.dataTransfer.files[0];
    if (
      droppedFile?.type ===
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    ) {
      setFile(droppedFile);
    } else {
      toast.error('Only .xlsx files are allowed!');
    }
  };

  const handleDragOver = (e) => {
    e.preventDefault();
  };

  const handleUpload = async (e) => {
    e.preventDefault();

    if (!file) {
      toast.error('Please select a file before uploading!');
      return;
    }

    const formData = new FormData();
    formData.append('file', file);
    setIsLoading(true);

    try {
      // === Option 1: Using fetch ===
      const response = await fetch('http://localhost:5000/upload', {
        method: 'POST',
        body: formData,
      });

      if (!response.ok) {
        const errorData = await response.json();
        toast.error(errorData?.error || 'Upload failed!');
        return;
      }

      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'Pubs_Updated.xlsx';
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(url);

      toast.success('File processed and downloaded successfully!');

      // === Option 2: Using axios ===
      /*
      const response = await axios.post('http://localhost:5000/upload', formData, {
        responseType: 'blob',
      });

      const blob = new Blob([response.data], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });

      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'Pubs_Updated.xlsx';
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(url);
      toast.success('File processed and downloaded successfully!');
      */
    } catch (error) {
      toast.error('Upload failed!');
    } finally {
      setIsLoading(false);
    }
  };

  return (
    <div className="min-h-screen flex items-center justify-center bg-gradient-to-br from-indigo-100 to-white p-4">
      <form
        onSubmit={handleUpload}
        className="bg-white rounded-2xl shadow-lg border border-gray-200 p-8 w-full max-w-md space-y-6"
      >
        <h2 className="text-2xl font-bold text-center text-gray-800">
          PubMed Excel Processor
        </h2>

        {/* Drag & Drop Area */}
        <div
          onDrop={handleDrop}
          onDragOver={handleDragOver}
          className="border-2 border-dashed border-indigo-400 p-6 rounded-md text-center bg-indigo-50 hover:bg-indigo-100 transition"
        >
          <p className="text-gray-700 text-sm">Drag & drop your .xlsx file here</p>
          <p className="text-gray-500 text-xs mt-1">or select a file below</p>
        </div>

        {/* File Input */}
        <input
          type="file"
          accept=".xlsx"
          onChange={(e) => setFile(e.target.files[0])}
          className="w-full border border-gray-300 rounded-md px-4 py-2 text-sm text-gray-700 mt-2 focus:outline-none focus:ring-2 focus:ring-indigo-400"
        />

        {/* File Preview */}
        {file && (
          <p className="text-sm text-green-600 font-medium truncate">
            Selected: {file.name}
          </p>
        )}

        {/* Upload Button */}
        <button
          type="submit"
          disabled={isLoading}
          className="w-full bg-indigo-600 text-white py-2 rounded-md hover:bg-indigo-700 transition disabled:opacity-50"
        >
          {isLoading ? 'Processing...' : 'Upload & Download'}
        </button>

        <p className="text-xs text-gray-400 text-center">Supported: .xlsx only</p>
      </form>

      <ToastContainer position="top-center" />
    </div>
  );
};

export default UploadForm;

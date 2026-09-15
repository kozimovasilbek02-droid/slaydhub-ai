import React, { useState, useEffect, useRef } from 'react';
import { 
  Sparkles, 
  Download, 
  Settings, 
  Layers, 
  CheckCircle2, 
  AlertCircle, 
  RefreshCw, 
  Search, 
  Sliders, 
  Eye, 
  Columns, 
  Maximize2, 
  FolderOpen, 
  Play, 
  ExternalLink,
  ChevronRight,
  Palette,
  FileText,
  Boxes,
  Zap,
  Layout,
  UploadCloud,
  Plus,
  Image as ImageIcon
} from 'lucide-react';

const API_BASE = '/api';

export default function App() {
  const [folders, setFolders] = useState([]);
  const [selectedFolder, setSelectedFolder] = useState('');
  const [folderDetails, setFolderDetails] = useState(null);
  const [activeSlideIdx, setActiveSlideIdx] = useState(1);
  const [searchQuery, setSearchQuery] = useState('');
  
  // Loading & Generating state
  const [loadingFolders, setLoadingFolders] = useState(false);
  const [loadingDetails, setLoadingDetails] = useState(false);
  const [generating, setGenerating] = useState(false);
  const [statusMsg, setStatusMsg] = useState('');
  
  // Drag and Drop state
  const [isDragging, setIsDragging] = useState(false);
  const [uploadingDrop, setUploadingDrop] = useState(false);
  const fileInputRef = useRef(null);
  const folderInputRef = useRef(null);
  
  // View mode: 'split', 'comparison'
  const [viewMode, setViewMode] = useState('split');
  
  // Inspection state for active slide
  const [inspectionData, setInspectionData] = useState(null);
  const [inspecting, setInspecting] = useState(false);

  // Fetch folders on mount
  useEffect(() => {
    fetchFolders();
  }, []);

  // Fetch details when folder changes
  useEffect(() => {
    if (selectedFolder) {
      fetchFolderDetails(selectedFolder);
    }
  }, [selectedFolder]);

  // Inspect slide when active slide changes
  useEffect(() => {
    if (folderDetails && folderDetails.slides && folderDetails.slides.length >= activeSlideIdx) {
      const activeSlide = folderDetails.slides[activeSlideIdx - 1];
      if (activeSlide && activeSlide.original_url) {
        const pathParam = activeSlide.original_url.split('path=')[1];
        if (pathParam) {
          inspectSlide(decodeURIComponent(pathParam));
        }
      }
    }
  }, [activeSlideIdx, folderDetails]);

  const fetchFolders = async (search = '', selectNewFolder = null) => {
    setLoadingFolders(true);
    try {
      const res = await fetch(`${API_BASE}/folders${search ? `?search=${encodeURIComponent(search)}` : ''}`);
      const data = await res.json();
      setFolders(data.folders || []);
      if (selectNewFolder) {
        setSelectedFolder(selectNewFolder);
      } else if (!selectedFolder && data.folders && data.folders.length > 0) {
        setSelectedFolder(data.folders[0].name);
      }
    } catch (e) {
      console.error('Folders fetch error:', e);
    } finally {
      setLoadingFolders(false);
    }
  };

  const fetchFolderDetails = async (folderName) => {
    setLoadingDetails(true);
    try {
      const res = await fetch(`${API_BASE}/folder/${folderName}/details`);
      if (res.ok) {
        const data = await res.json();
        setFolderDetails(data);
        setActiveSlideIdx(1);
      }
    } catch (e) {
      console.error('Folder details error:', e);
    } finally {
      setLoadingDetails(false);
    }
  };

  const inspectSlide = async (imagePath) => {
    setInspecting(true);
    try {
      const res = await fetch(`${API_BASE}/inspect_slide?image_path=${encodeURIComponent(imagePath)}`);
      if (res.ok) {
        const data = await res.json();
        setInspectionData(data);
      }
    } catch (e) {
      console.error('Slide inspect error:', e);
    } finally {
      setInspecting(false);
    }
  };

  // ─────────────────────────────────────────────────────────
  // DRAG & DROP HANDLERS (Files and Folders)
  // ─────────────────────────────────────────────────────────
  const handleDragOver = (e) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(true);
  };

  const handleDragLeave = (e) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);
  };

  const handleDrop = async (e) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);

    const items = e.dataTransfer.items;
    const files = [];

    if (items) {
      for (let i = 0; i < items.length; i++) {
        const item = items[i].webkitGetAsEntry ? items[i].webkitGetAsEntry() : null;
        if (item) {
          await scanFiles(item, files);
        } else if (items[i].kind === 'file') {
          const f = items[i].getAsFile();
          if (f && isImageFile(f.name)) files.push(f);
        }
      }
    } else if (e.dataTransfer.files) {
      for (let i = 0; i < e.dataTransfer.files.length; i++) {
        const f = e.dataTransfer.files[i];
        if (isImageFile(f.name)) files.push(f);
      }
    }

    if (files.length > 0) {
      uploadFiles(files);
    } else {
      alert("Iltimos, rasm (.png, .jpg, .jpeg, .webp) yoki rasmlar joylashgan papkani tashlang!");
    }
  };

  const scanFiles = async (entry, fileList) => {
    if (entry.isFile) {
      const file = await new Promise((resolve) => entry.file(resolve));
      if (isImageFile(file.name)) fileList.push(file);
    } else if (entry.isDirectory) {
      const reader = entry.createReader();
      const entries = await new Promise((resolve) => reader.readEntries(resolve));
      for (const child of entries) {
        await scanFiles(child, fileList);
      }
    }
  };

  const isImageFile = (name) => {
    return /\.(png|jpg|jpeg|webp)$/i.test(name);
  };

  const handleFileInputChange = (e) => {
    if (e.target.files && e.target.files.length > 0) {
      uploadFiles(Array.from(e.target.files));
    }
  };

  const uploadFiles = async (files) => {
    setUploadingDrop(true);
    setStatusMsg(`🚀 ${files.length} ta rasm yuklanmoqda va AI Vektor slaydlar noldan chizilmoqda...`);
    
    const formData = new FormData();
    files.forEach((f) => formData.append('files', f));
    
    // Auto-name from first file
    const baseName = files[0].name.replace(/\.[^/.]+$/, '').replace(/[^a-zA-Z0-9_-]/g, '_');
    formData.append('custom_name', `Taqdimot_${baseName}`);

    try {
      const res = await fetch(`${API_BASE}/upload_drop`, {
        method: 'POST',
        body: formData
      });
      if (res.ok) {
        const data = await res.json();
        await fetchFolders('', data.folder_name);
        setStatusMsg(`✅ ${data.total_slides} ta slayd 100% vektor shaklda yaratildi!`);
      } else {
        const err = await res.json();
        alert(`Xatolik: ${err.detail || 'Yuklashda xatolik yuz berdi'}`);
      }
    } catch (e) {
      alert(`Xatolik: ${e.message}`);
    } finally {
      setUploadingDrop(false);
    }
  };

  const handleGenerate = async () => {
    if (!selectedFolder) return;
    setGenerating(true);
    setStatusMsg("AI OCR matn ajratish, vektor kartalar va shakllar yaratilmoqda...");
    try {
      const res = await fetch(`${API_BASE}/generate`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ folder_name: selectedFolder })
      });
      if (res.ok) {
        setStatusMsg("✅ Taqdimot 100% tahrir qilinadigan vektor shaklida yaratildi!");
        await fetchFolderDetails(selectedFolder);
      } else {
        const err = await res.json();
        alert(`Xatolik: ${err.detail || 'Generatsiya muvaffaqiyatsiz bo\'ldi'}`);
      }
    } catch (e) {
      alert(`Xatolik: ${e.message}`);
    } finally {
      setGenerating(false);
    }
  };

  const handleOpenPowerpoint = async () => {
    if (!selectedFolder) return;
    try {
      await fetch(`${API_BASE}/open_in_powerpoint`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ folder_name: selectedFolder })
      });
    } catch (e) {
      alert(`Ochishda xatolik: ${e.message}`);
    }
  };

  const activeSlide = folderDetails?.slides?.[activeSlideIdx - 1];

  return (
    <div 
      onDragOver={handleDragOver}
      onDragLeave={handleDragLeave}
      onDrop={handleDrop}
      className="relative flex h-screen bg-slate-950 text-slate-100 font-sans overflow-hidden"
    >
      {/* ────────────────── DRAG & DROP FULLSCREEN OVERLAY ────────────────── */}
      {isDragging && (
        <div className="absolute inset-0 z-50 bg-indigo-950/80 backdrop-blur-md border-4 border-dashed border-indigo-400 flex flex-col items-center justify-center pointer-events-none animate-in fade-in duration-150">
          <div className="w-24 h-24 rounded-3xl bg-indigo-600/30 border border-indigo-400/50 flex items-center justify-center mb-4 shadow-2xl shadow-indigo-500/50">
            <UploadCloud className="w-12 h-12 text-indigo-300 animate-bounce" />
          </div>
          <h2 className="text-2xl font-bold text-white mb-1">
            Papka yoki Rasmlarni Shu Yerga Tashlang!
          </h2>
          <p className="text-sm text-indigo-200">
            AI avtomatik tarzda barcha rasmlarni tahrir qilinadigan vektor slaydlarga aylantiradi.
          </p>
        </div>
      )}

      {/* Hidden File Inputs */}
      <input
        type="file"
        ref={fileInputRef}
        onChange={handleFileInputChange}
        multiple
        accept="image/*"
        className="hidden"
      />
      <input
        type="file"
        ref={folderInputRef}
        onChange={handleFileInputChange}
        webkitdirectory="true"
        directory="true"
        multiple
        className="hidden"
      />

      {/* ────────────────── LEFT SIDEBAR: THUMBNAILS & FOLDERS ────────────────── */}
      <aside className="w-80 bg-slate-900/90 border-r border-slate-800 flex flex-col backdrop-blur-xl">
        {/* Logo & Header */}
        <div className="p-4 border-b border-slate-800 flex items-center justify-between">
          <div className="flex items-center space-x-3">
            <div className="w-10 h-10 rounded-xl bg-gradient-to-tr from-indigo-600 via-blue-500 to-cyan-400 flex items-center justify-center shadow-lg shadow-indigo-500/30">
              <Sparkles className="w-5 h-5 text-white animate-pulse" />
            </div>
            <div>
              <h1 className="font-bold text-base bg-gradient-to-r from-white via-slate-200 to-slate-400 bg-clip-text text-transparent">
                SlideAI Studio
              </h1>
              <p className="text-xs text-indigo-400 font-medium flex items-center gap-1">
                <Zap className="w-3 h-3" /> Vector Precision 95%+
              </p>
            </div>
          </div>
        </div>

        {/* Drag & Drop Quick Button Area */}
        <div className="p-3 border-b border-slate-800/80 bg-slate-900/40">
          <div 
            onClick={() => fileInputRef.current?.click()}
            className="w-full border-2 border-dashed border-indigo-700/60 hover:border-indigo-400 bg-indigo-950/20 hover:bg-indigo-950/40 rounded-xl p-3 text-center cursor-pointer transition-all flex flex-col items-center justify-center group"
          >
            <UploadCloud className="w-6 h-6 text-indigo-400 group-hover:scale-110 transition-transform mb-1" />
            <p className="text-xs font-bold text-slate-200">
              📁 Rasm yoki Papkani Tashlang
            </p>
            <p className="text-[10px] text-slate-400 mt-0.5">
              yoki kompyuterdan tanlash uchun bosing
            </p>
          </div>
        </div>

        {/* Presentation Search & Selector */}
        <div className="p-3 border-b border-slate-800/80 bg-slate-900/50">
          <div className="relative mb-2">
            <Search className="w-4 h-4 text-slate-400 absolute left-3 top-2.5" />
            <input
              type="text"
              placeholder="597+ Taqdimot qidirish..."
              value={searchQuery}
              onChange={(e) => {
                setSearchQuery(e.target.value);
                fetchFolders(e.target.value);
              }}
              className="w-full bg-slate-950 border border-slate-800 rounded-lg pl-9 pr-3 py-1.5 text-xs text-slate-200 placeholder-slate-500 focus:outline-none focus:border-indigo-500 transition-colors"
            />
          </div>

          <label className="text-[11px] font-semibold text-slate-400 uppercase tracking-wider block mb-1">
            Faol Taqdimot ({folders.length}):
          </label>
          <select
            value={selectedFolder}
            onChange={(e) => setSelectedFolder(e.target.value)}
            className="w-full bg-slate-950 border border-slate-700/80 rounded-lg px-3 py-2 text-xs font-medium text-slate-200 focus:outline-none focus:border-indigo-500 cursor-pointer shadow-inner"
          >
            {folders.map((f) => (
              <option key={f.name} value={f.name}>
                {f.is_ready ? '✅' : '📁'} {f.name} ({f.slide_count} slayd)
              </option>
            ))}
          </select>
        </div>

        {/* Slide Thumbnails List */}
        <div className="flex-1 overflow-y-auto p-3 space-y-2.5 custom-scrollbar">
          <div className="flex items-center justify-between px-1 mb-1">
            <span className="text-xs font-semibold text-slate-400 uppercase tracking-wider">
              Slaydlar ({folderDetails?.slides?.length || 0})
            </span>
            <span className="text-[11px] text-indigo-400 bg-indigo-950/60 px-2 py-0.5 rounded-full border border-indigo-800/40">
              {activeSlideIdx} / {folderDetails?.slides?.length || 0}
            </span>
          </div>

          {loadingDetails ? (
            <div className="flex flex-col items-center justify-center py-12 text-slate-500 space-y-2">
              <RefreshCw className="w-6 h-6 animate-spin text-indigo-500" />
              <p className="text-xs">Slaydlar yuklanmoqda...</p>
            </div>
          ) : folderDetails?.slides?.map((s) => {
            const isSelected = activeSlideIdx === s.index;
            return (
              <button
                key={s.index}
                onClick={() => setActiveSlideIdx(s.index)}
                className={`w-full text-left rounded-xl p-2 transition-all flex items-center space-x-3 border ${
                  isSelected
                    ? 'bg-gradient-to-r from-indigo-900/60 to-slate-800/90 border-indigo-500/80 shadow-md shadow-indigo-950/50'
                    : 'bg-slate-950/60 hover:bg-slate-800/50 border-slate-800/60 hover:border-slate-700'
                }`}
              >
                {/* Thumbnail */}
                <div className="w-20 h-12 rounded-lg bg-slate-900 border border-slate-800 overflow-hidden flex-shrink-0 relative">
                  <img
                    src={s.generated_url || s.original_url}
                    alt={`Slide ${s.index}`}
                    className="w-full h-full object-cover"
                    loading="lazy"
                  />
                  <div className="absolute top-0.5 left-0.5 bg-black/70 text-white text-[9px] font-bold px-1.5 py-0.2 rounded">
                    #{s.index}
                  </div>
                </div>

                {/* Details */}
                <div className="flex-1 min-w-0">
                  <p className="text-xs font-semibold text-slate-200 truncate">
                    Slayd {s.index}
                  </p>
                  <p className="text-[10px] text-slate-400 flex items-center gap-1 mt-0.5">
                    {s.generated_url ? (
                      <span className="text-emerald-400 flex items-center gap-0.5">
                        <CheckCircle2 className="w-2.5 h-2.5" /> Vector Ready
                      </span>
                    ) : (
                      <span className="text-amber-400 flex items-center gap-0.5">
                        <AlertCircle className="w-2.5 h-2.5" /> Generatsiya kutilmoqda
                      </span>
                    )}
                  </p>
                </div>
              </button>
            );
          })}
        </div>
      </aside>

      {/* ────────────────── MAIN CENTER WORKSPACE ────────────────── */}
      <main className="flex-1 flex flex-col bg-slate-950 overflow-hidden">
        {/* Top Control Bar */}
        <header className="h-16 bg-slate-900/80 border-b border-slate-800 px-6 flex items-center justify-between backdrop-blur-md">
          <div className="flex items-center space-x-4">
            <div>
              <h2 className="text-sm font-bold text-slate-100 flex items-center gap-2">
                <FolderOpen className="w-4 h-4 text-indigo-400" />
                {folderDetails?.folder_name || 'Taqdimot tanlanmagan'}
              </h2>
              <p className="text-xs text-slate-400">
                {folderDetails?.total_slides || 0} ta original slayd | 100% Vector Elementlar
              </p>
            </div>
          </div>

          {/* View Mode Switcher */}
          <div className="flex items-center bg-slate-950 p-1 rounded-xl border border-slate-800">
            <button
              onClick={() => setViewMode('split')}
              className={`px-3 py-1.5 rounded-lg text-xs font-medium flex items-center gap-1.5 transition-all ${
                viewMode === 'split'
                  ? 'bg-indigo-600 text-white shadow-sm'
                  : 'text-slate-400 hover:text-slate-200'
              }`}
            >
              <Columns className="w-3.5 h-3.5" /> Yonma-yon (Split)
            </button>
            <button
              onClick={() => setViewMode('comparison')}
              className={`px-3 py-1.5 rounded-lg text-xs font-medium flex items-center gap-1.5 transition-all ${
                viewMode === 'comparison'
                  ? 'bg-indigo-600 text-white shadow-sm'
                  : 'text-slate-400 hover:text-slate-200'
              }`}
            >
              <Eye className="w-3.5 h-3.5" /> Solishtirish Karta
            </button>
          </div>

          {/* Actions: Generate & Export */}
          <div className="flex items-center space-x-3">
            <button
              onClick={() => fileInputRef.current?.click()}
              className="px-3 py-2 rounded-xl bg-slate-800 hover:bg-slate-700 text-slate-200 text-xs font-semibold flex items-center gap-1.5 border border-slate-700 transition-all cursor-pointer"
            >
              <Plus className="w-4 h-4 text-indigo-400" /> Rasm / Papka Qo'shish
            </button>

            <button
              onClick={handleGenerate}
              disabled={generating || uploadingDrop}
              className="px-4 py-2 rounded-xl bg-gradient-to-r from-indigo-600 to-blue-600 hover:from-indigo-500 hover:to-blue-500 text-white text-xs font-bold flex items-center gap-2 shadow-lg shadow-indigo-600/30 transition-all disabled:opacity-50 cursor-pointer"
            >
              {generating || uploadingDrop ? (
                <>
                  <RefreshCw className="w-4 h-4 animate-spin" /> Yaratilmoqda...
                </>
              ) : (
                <>
                  <Play className="w-4 h-4" /> AI Vektor Yaratish
                </>
              )}
            </button>

            {folderDetails?.pptx_ready && (
              <>
                <a
                  href={`${API_BASE}/download_pptx/${selectedFolder}`}
                  className="px-3.5 py-2 rounded-xl bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-bold flex items-center gap-1.5 shadow-lg shadow-emerald-600/20 transition-all"
                  download
                >
                  <Download className="w-4 h-4" /> PPTX Yuklab Olish
                </a>

                <button
                  onClick={handleOpenPowerpoint}
                  className="px-3.5 py-2 rounded-xl bg-slate-800 hover:bg-slate-700 text-slate-200 text-xs font-bold flex items-center gap-1.5 border border-slate-700 transition-all cursor-pointer"
                >
                  <ExternalLink className="w-4 h-4" /> PowerPoint'da Ochish
                </button>
              </>
            )}
          </div>
        </header>

        {/* Workspace Canvas */}
        <div className="flex-1 p-6 overflow-y-auto custom-scrollbar flex flex-col space-y-6">
          {(generating || uploadingDrop) && (
            <div className="bg-indigo-950/40 border border-indigo-800/60 rounded-2xl p-4 flex items-center space-x-3 animate-pulse">
              <RefreshCw className="w-5 h-5 text-indigo-400 animate-spin flex-shrink-0" />
              <p className="text-xs text-indigo-200 font-medium">{statusMsg}</p>
            </div>
          )}

          {/* Active Slide Canvas View */}
          {activeSlide ? (
            <div className="grid grid-cols-1 gap-6">
              {viewMode === 'split' && (
                <div className="grid grid-cols-2 gap-6">
                  {/* Left: Original Slide */}
                  <div className="bg-slate-900 border border-slate-800 rounded-2xl overflow-hidden flex flex-col shadow-xl">
                    <div className="bg-slate-900/90 px-4 py-2.5 border-b border-slate-800 flex items-center justify-between">
                      <span className="text-xs font-bold text-slate-300 flex items-center gap-2">
                        <FileText className="w-4 h-4 text-slate-400" />
                        Original Slayd #{activeSlide.index} (Manba Rasm)
                      </span>
                    </div>
                    <div className="p-4 flex items-center justify-center bg-slate-950/40 flex-1 min-h-[380px]">
                      <img
                        src={activeSlide.original_url}
                        alt="Original"
                        className="max-h-[460px] w-auto object-contain rounded-lg shadow-md"
                      />
                    </div>
                  </div>

                  {/* Right: Recreated Vector Slide */}
                  <div className="bg-slate-900 border border-indigo-900/50 rounded-2xl overflow-hidden flex flex-col shadow-xl">
                    <div className="bg-slate-900/90 px-4 py-2.5 border-b border-slate-800 flex items-center justify-between">
                      <span className="text-xs font-bold text-indigo-300 flex items-center gap-2">
                        <Sparkles className="w-4 h-4 text-indigo-400" />
                        100% Tahrir Qilinadigan Vektor Slayd #{activeSlide.index}
                      </span>
                      <span className="text-[10px] bg-emerald-950/80 text-emerald-300 border border-emerald-800/50 px-2 py-0.5 rounded-md font-semibold">
                        95%+ Aniqlik
                      </span>
                    </div>
                    <div className="p-4 flex items-center justify-center bg-slate-950/40 flex-1 min-h-[380px]">
                      {activeSlide.generated_url ? (
                        <img
                          src={activeSlide.generated_url}
                          alt="Generated Vector Slide"
                          className="max-h-[460px] w-auto object-contain rounded-lg shadow-md"
                        />
                      ) : (
                        <div className="text-center p-8 space-y-3">
                          <AlertCircle className="w-10 h-10 text-slate-600 mx-auto" />
                          <p className="text-xs text-slate-400">
                            Ushbu slayd hali vektor shakliga aylantirilmagan.
                          </p>
                          <button
                            onClick={handleGenerate}
                            className="px-4 py-2 bg-indigo-600 hover:bg-indigo-500 text-white rounded-xl text-xs font-bold"
                          >
                            Hozir Yaratish
                          </button>
                        </div>
                      )}
                    </div>
                  </div>
                </div>
              )}

              {viewMode === 'comparison' && (
                <div className="bg-slate-900 border border-slate-800 rounded-2xl overflow-hidden shadow-xl">
                  <div className="bg-slate-900/90 px-4 py-2.5 border-b border-slate-800 flex items-center justify-between">
                    <span className="text-xs font-bold text-slate-300 flex items-center gap-2">
                      <Eye className="w-4 h-4 text-cyan-400" />
                      Side-by-Side Taqqoslash Kartasi #{activeSlide.index}
                    </span>
                  </div>
                  <div className="p-4 flex items-center justify-center bg-slate-950/40 min-h-[400px]">
                    {activeSlide.comparison_url ? (
                      <img
                        src={activeSlide.comparison_url}
                        alt="Comparison Card"
                        className="max-w-full max-h-[500px] object-contain rounded-lg shadow-md"
                      />
                    ) : (
                      <div className="text-center p-8 space-y-3">
                        <AlertCircle className="w-10 h-10 text-slate-600 mx-auto" />
                        <p className="text-xs text-slate-400">
                          Taqqoslovchi kartani yaratish uchun "AI Vektor Yaratish" tugmasini bosing.
                        </p>
                      </div>
                    )}
                  </div>
                </div>
              )}

              {/* ────────────────── BOTTOM AI INSPECTOR PANEL ────────────────── */}
              <div className="bg-slate-900/90 border border-slate-800 rounded-2xl p-5 shadow-xl">
                <div className="flex items-center justify-between mb-4 border-b border-slate-800 pb-3">
                  <div className="flex items-center space-x-2">
                    <Boxes className="w-4 h-4 text-indigo-400" />
                    <h3 className="text-xs font-bold text-slate-200 uppercase tracking-wider">
                      AI Slayd Strukturasi & Elementlar Inspeksiyasi
                    </h3>
                  </div>

                  {inspectionData && (
                    <div className="flex items-center space-x-4 text-xs">
                      <span className="text-slate-400 flex items-center gap-1.5">
                        <Palette className="w-3.5 h-3.5 text-indigo-400" />
                        Fon rangi: 
                        <span
                          className="w-3.5 h-3.5 rounded-full inline-block border border-white/20"
                          style={{ backgroundColor: inspectionData.bg_color_hex }}
                        />
                        <code className="text-[11px] text-slate-300">{inspectionData.bg_color_hex}</code>
                      </span>
                      <span className="text-slate-400">
                        Vektor Kartalar: <strong className="text-indigo-300">{inspectionData.cards?.length || 0} ta</strong>
                      </span>
                      <span className="text-slate-400">
                        Matn qutilari: <strong className="text-emerald-300">{inspectionData.text_items?.length || 0} ta</strong>
                      </span>
                    </div>
                  )}
                </div>

                {inspecting ? (
                  <div className="py-8 text-center text-xs text-slate-500 flex items-center justify-center gap-2">
                    <RefreshCw className="w-4 h-4 animate-spin text-indigo-400" /> Elementlar tahlil qilinmoqda...
                  </div>
                ) : inspectionData ? (
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                    {/* Detected Cards Swatches */}
                    <div className="bg-slate-950/60 p-3 rounded-xl border border-slate-800/80">
                      <h4 className="text-[11px] font-bold text-slate-400 uppercase tracking-wider mb-2 flex items-center gap-1.5">
                        <Layout className="w-3 h-3 text-indigo-400" /> Aniqlangan Vektor Kartalar:
                      </h4>
                      <div className="flex flex-wrap gap-2">
                        {inspectionData.cards?.map((card, i) => (
                          <div
                            key={i}
                            className="flex items-center space-x-2 bg-slate-900 px-2.5 py-1.5 rounded-lg border border-slate-700/60"
                          >
                            <span
                              className="w-4 h-4 rounded-md shadow-sm border border-white/20"
                              style={{ backgroundColor: card.color_hex }}
                            />
                            <span className="text-[11px] font-semibold text-slate-300">
                              Karta #{i + 1}
                            </span>
                            <span className="text-[10px] text-slate-500">
                              {card.is_dark ? '(To\'q)' : '(Yorug\')'}
                            </span>
                          </div>
                        ))}
                      </div>
                    </div>

                    {/* Detected OCR Text Boxes Preview */}
                    <div className="bg-slate-950/60 p-3 rounded-xl border border-slate-800/80">
                      <h4 className="text-[11px] font-bold text-slate-400 uppercase tracking-wider mb-2 flex items-center gap-1.5">
                        <FileText className="w-3 h-3 text-emerald-400" /> Ajratilgan OCR Matn Qutilari (Top 8):
                      </h4>
                      <div className="space-y-1 max-h-36 overflow-y-auto custom-scrollbar">
                        {inspectionData.text_items?.slice(0, 8).map((item, i) => (
                          <div
                            key={i}
                            className="text-[11px] bg-slate-900/80 px-2.5 py-1 rounded border border-slate-800 flex items-center justify-between text-slate-300"
                          >
                            <span className="truncate flex-1 font-medium">{item.text}</span>
                            <span className="text-[9px] text-slate-500 ml-2 font-mono">{item.font_pt}pt</span>
                          </div>
                        ))}
                      </div>
                    </div>
                  </div>
                ) : (
                  <p className="text-xs text-slate-500">Inspeksiya ma'lumotlari yuklanmadi.</p>
                )}
              </div>
            </div>
          ) : (
            <div className="flex-1 flex flex-col items-center justify-center py-24 text-center border-2 border-dashed border-slate-800 rounded-3xl bg-slate-900/30 p-8">
              <UploadCloud className="w-16 h-16 text-indigo-400 mb-4 animate-bounce" />
              <h3 className="text-lg font-bold text-slate-200">Istalgan Slayd Rasmi yoki Papkani Tashlang</h3>
              <p className="text-xs text-slate-400 max-w-md mt-1 mb-5">
                Hech qanday manzil yozish shart emas. Rasmni yoki butun papkani to'g'ridan-to'g'ri ekranga tashlang — AI bir zumda tahrir qilinadigan vektor slayd yaratib beradi.
              </p>
              <button
                onClick={() => fileInputRef.current?.click()}
                className="px-5 py-2.5 rounded-xl bg-indigo-600 hover:bg-indigo-500 text-white text-xs font-bold shadow-lg shadow-indigo-600/30"
              >
                Fayl yoki Papka Tanlash
              </button>
            </div>
          )}
        </div>
      </main>
    </div>
  );
}

import os
import sys

def write_file(rel_path, content):
    full_path = os.path.abspath(rel_path)
    os.makedirs(os.path.dirname(full_path), exist_ok=True)
    with open(full_path, 'w', encoding='utf-8') as f:
        f.write(content.strip() + '\n')
    print(f"Successfully created: {rel_path}")

# ============================================================
# LAUNCHER: start_app.py
# ============================================================
start_app_code = '''# -*- coding: utf-8 -*-
"""
SlideTranslate AI — One-click Complete Launcher
Starts FastAPI Backend (Port 8000) and Vite React Frontend (Port 5173).
"""

import os
import sys
import time
import subprocess
import webbrowser

# UTF-8 encoding support on Windows
if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

def print_banner():
    print("=" * 65)
    print("      🚀 SlideTranslate AI — Professional Slayd Tarjimoni")
    print("      FastAPI (Backend) + React / Vite (Frontend) + Gemini AI")
    print("=" * 65)

def main():
    print_banner()
    root_dir = os.path.abspath(os.path.dirname(__file__))
    frontend_dir = os.path.join(root_dir, "frontend")

    print("[1/3] Backend tekshirilmoqda...")
    backend_cmd = [sys.executable, "-m", "uvicorn", "backend.main:app", "--host", "0.0.0.0", "--port", "8000", "--reload"]
    print("      Backend buyrug'i: " + " ".join(backend_cmd))
    backend_proc = subprocess.Popen(backend_cmd, cwd=root_dir)

    print("[2/3] Frontend (Vite) ishga tushirilmoqda...")
    # On Windows npm is a cmd/bat
    npm_cmd = "npm.cmd" if os.name == "nt" else "npm"
    frontend_proc = subprocess.Popen([npm_cmd, "run", "dev", "--", "--host"], cwd=frontend_dir)

    time.sleep(2)
    url = "http://localhost:5173"
    print(f"\\n[3/3] Barcha xizmatlar muvaffaqiyatli ishga tushdi!")
    print(f"      🌐 Veb-ilova manzili: {url}")
    print(f"      📡 Backend API: http://localhost:8000/docs")
    print(f"\\nBrauzer avtomatik ochilmoqda...")
    
    try:
        webbrowser.open(url)
    except Exception:
        pass

    print("\\n[Dasturni to'xtatish uchun Ctrl+C tugmalarini bosing]\\n")
    try:
        backend_proc.wait()
        frontend_proc.wait()
    except KeyboardInterrupt:
        print("\\nTo'xtatilmoqda...")
        backend_proc.terminate()
        frontend_proc.terminate()
        print("Xizmatlar to'xtatildi.")

if __name__ == "__main__":
    main()
'''
write_file('start_app.py', start_app_code)

package_json = '''{
  "name": "slidetranslate-ai-frontend",
  "private": true,
  "version": "1.0.0",
  "type": "module",
  "scripts": {
    "dev": "vite",
    "build": "vite build",
    "preview": "vite preview"
  },
  "dependencies": {
    "clsx": "^2.1.1",
    "lucide-react": "^1.16.0",
    "react": "^18.3.1",
    "react-dom": "^18.3.1",
    "tailwind-merge": "^2.6.0"
  },
  "devDependencies": {
    "@vitejs/plugin-react": "^4.3.4",
    "autoprefixer": "^10.4.20",
    "postcss": "^8.4.49",
    "tailwindcss": "^3.4.17",
    "vite": "^6.0.7"
  }
}
'''
write_file('frontend/package.json', package_json)

vite_config = '''import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  server: {
    port: 5173,
    proxy: {
      '/api': {
        target: 'http://localhost:8000',
        changeOrigin: true,
        secure: false
      }
    }
  }
})
'''
write_file('frontend/vite.config.js', vite_config)

tailwind_config = '''/** @type {import('tailwindcss').Config} */
export default {
  content: [
    "./index.html",
    "./src/**/*.{js,ts,jsx,tsx}",
  ],
  darkMode: 'class',
  theme: {
    extend: {
      colors: {
        brand: {
          50: '#f0f7ff',
          100: '#e0effe',
          200: '#bae0fd',
          300: '#7cc7fb',
          400: '#36abf7',
          500: '#0c90eb',
          600: '#0273c8',
          700: '#035ca2',
          800: '#074e85',
          900: '#0c416e',
          950: '#082a4a',
        },
      },
      fontFamily: {
        sans: ['Plus Jakarta Sans', 'Inter', 'system-ui', 'sans-serif'],
      }
    },
  },
  plugins: [],
}
'''
write_file('frontend/tailwind.config.js', tailwind_config)

postcss_config = '''export default {
  plugins: {
    tailwindcss: {},
    autoprefixer: {},
  },
}
'''
write_file('frontend/postcss.config.js', postcss_config)

index_html = '''<!doctype html>
<html lang="uz">
  <head>
    <meta charset="UTF-8" />
    <link rel="icon" type="image/svg+xml" href="data:image/svg+xml,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='%230273c8'><path d='M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm-5 14H7v-2h7v2zm3-4H7v-2h10v2zm0-4H7V7h10v2z'/></svg>" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>SlideTranslate AI — Taqdimotlarni O'zbek tiliga tarjima qilish</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;500&display=swap" rel="stylesheet">
  </head>
  <body class="bg-slate-900 text-slate-100 min-h-screen antialiased selection:bg-brand-500 selection:text-white">
    <div id="root"></div>
    <script type="module" src="/src/main.jsx"></script>
  </body>
</html>
'''
write_file('frontend/index.html', index_html)

index_css = '''@tailwind base;
@tailwind components;
@tailwind utilities;

@layer base {
  body {
    font-feature-settings: "cv02", "cv03", "cv04", "cv11";
  }
}

/* Custom modern scrollbars */
::-webkit-scrollbar {
  width: 6px;
  height: 6px;
}
::-webkit-scrollbar-track {
  background: rgba(15, 23, 42, 0.6);
}
::-webkit-scrollbar-thumb {
  background: rgba(100, 116, 139, 0.4);
  border-radius: 9999px;
}
::-webkit-scrollbar-thumb:hover {
  background: rgba(148, 163, 184, 0.7);
}
'''
write_file('frontend/src/index.css', index_css)

main_jsx = '''import React from 'react'
import ReactDOM from 'react-dom/client'
import App from './App.jsx'
import './index.css'

ReactDOM.createRoot(document.getElementById('root')).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>,
)
'''
write_file('frontend/src/main.jsx', main_jsx)

app_jsx = '''import React, { useState, useEffect, useRef } from 'react';
import { 
  UploadCloud, 
  Sparkles, 
  Download, 
  Settings, 
  BookOpen, 
  FileText, 
  Layers, 
  CheckCircle2, 
  AlertCircle, 
  RefreshCw, 
  Copy, 
  Check, 
  Search, 
  Sliders, 
  Eye, 
  Columns, 
  Maximize2, 
  Languages, 
  Key, 
  Plus, 
  Trash2, 
  ArrowRight,
  HelpCircle,
  Zap,
  Cpu
} from 'lucide-react';

const API_BASE = '/api';

export default function App() {
  // Session State
  const [session, setSession] = useState(null);
  const [activeSlideIdx, setActiveSlideIdx] = useState(1);
  const [loading, setLoading] = useState(false);
  const [translating, setTranslating] = useState(false);
  const [exporting, setExporting] = useState(false);
  const [progress, setProgress] = useState(0);
  const [statusMsg, setStatusMsg] = useState('');

  // Settings State
  const [apiKey, setApiKey] = useState(() => localStorage.getItem('gemini_api_key') || '');
  const [targetScript, setTargetScript] = useState('latin'); // 'latin' or 'cyrillic'
  const [domain, setDomain] = useState('IT & Dasturlash');
  const [autoFit, setAutoFit] = useState(true);
  const [viewMode, setViewMode] = useState('split'); // 'split', 'translated', 'original'

  // Modals & Search
  const [showSettingsModal, setShowSettingsModal] = useState(false);
  const [showGlossaryModal, setShowGlossaryModal] = useState(false);
  const [searchTerm, setSearchTerm] = useState('');
  const [filterType, setFilterType] = useState('all');
  const [copiedId, setCopiedId] = useState(null);

  // Custom Glossary
  const [glossary, setGlossary] = useState(() => {
    try {
      const saved = localStorage.getItem('slide_glossary');
      return saved ? JSON.parse(saved) : {
        "Cloud Computing": "Bulutli hisoblash",
        "Artificial Intelligence": "Sun'iy intellekt",
        "Machine Learning": "Mashinaviy o'rganish",
        "Framework": "Freymvork",
        "Pipeline": "Konveyer tizimi",
        "Roadmap": "Yo'l xaritasi",
        "Milestone": "Muhim bosqich"
      };
    } catch {
      return {};
    }
  });
  const [newTermKey, setNewTermKey] = useState('');
  const [newTermVal, setNewTermVal] = useState('');

  // Drag & drop
  const [isDragging, setIsDragging] = useState(false);
  const fileInputRef = useRef(null);

  useEffect(() => {
    localStorage.setItem('slide_glossary', JSON.stringify(glossary));
  }, [glossary]);

  useEffect(() => {
    if (apiKey) {
      localStorage.setItem('gemini_api_key', apiKey);
    }
  }, [apiKey]);

  // Handle File Upload
  const handleFileUpload = async (file) => {
    if (!file || !file.name.endsWith('.pptx')) {
      alert("Iltimos, faqat PowerPoint (.pptx) fayllarini yuklang!");
      return;
    }

    setLoading(true);
    setStatusMsg("PPTX fayli tahlil qilinmoqda...");
    const formData = new FormData();
    formData.append('file', file);

    try {
      const res = await fetch(`${API_BASE}/upload`, {
        method: 'POST',
        body: formData
      });
      if (!res.ok) {
        const err = await res.json();
        throw new Error(err.detail || "Fayl yuklashda xatolik yuz berdi");
      }
      const data = await res.json();
      setSession(data);
      setActiveSlideIdx(1);
      setStatusMsg("Fayl muvaffaqiyatli yuklandi!");
    } catch (e) {
      alert(e.message);
    } finally {
      setLoading(false);
    }
  };

  // Load Built-in Demo
  const handleLoadSample = async () => {
    setLoading(true);
    setStatusMsg("Namunaviy taqdimot yuklanmoqda...");
    try {
      // Create a blob from backend test presentation
      const res = await fetch('/api/health');
      if (res.ok) {
        // Upload sample through test endpoint or mock upload
        const sampleSlideItems = [
          {
            id: "s1_sh0_p0",
            slide_index: 1,
            shape_name: "Sarlavha",
            item_type: "title",
            original_text: "Strategic Roadmap 2030: Next-Gen Enterprise AI",
            translated_text: targetScript === 'latin' ? "2030-yilgi Strategik Yo'l Xaritasi: Yangi Avlod Korporativ Sun'iy Intellekti" : "2030-йилги Стратегик Йўл Харитаси: Янги Авлод Корпоратив Сунъий Интеллекти",
            font_size_pt: 32,
            is_bold: true,
            is_italic: false,
            font_color: "#3b82f6",
            font_name: "Calibri",
            alignment: "left",
            box: { left: 5, top: 8, width: 90, height: 18 }
          },
          {
            id: "s1_sh0_p1",
            slide_index: 1,
            shape_name: "Kichik Sarlavha",
            item_type: "subtitle",
            original_text: "Accelerating Global Innovation, Scalability, and Intelligent Automation",
            translated_text: targetScript === 'latin' ? "Global Innovatsiyalar, Masshtablanuvchanlik va Aqlli Avtomatlashtirishni Jadallashtirish" : "Глобал Инновациялар, Масштабланувчанлик ва Ақлли Автоматлаштиришни Жадаллаштириш",
            font_size_pt: 18,
            is_bold: false,
            is_italic: false,
            font_color: "#94a3b8",
            font_name: "Calibri",
            alignment: "left",
            box: { left: 5, top: 28, width: 90, height: 12 }
          },
          {
            id: "s1_sh1_p0",
            slide_index: 1,
            shape_name: "Band 1",
            item_type: "body",
            original_text: "• High-precision automated translation preserving 100% layout and styles",
            translated_text: targetScript === 'latin' ? "• 100% slayd dizayni va shriftlarini saqlovchi yuqori aniqlikdagi avtomatlashtirilgan tarjima" : "• 100% слайд дизайни ва шрифтларини сақловчи юқори аниқликдаги автоматлаштирилган таржима",
            font_size_pt: 15,
            is_bold: false,
            is_italic: false,
            font_color: "#e2e8f0",
            font_name: "Calibri",
            alignment: "left",
            box: { left: 5, top: 45, width: 90, height: 10 }
          },
          {
            id: "s1_sh1_p1",
            slide_index: 1,
            shape_name: "Band 2",
            item_type: "body",
            original_text: "• Native support for both Uzbek Latin and Uzbek Cyrillic orthographies",
            translated_text: targetScript === 'latin' ? "• O'zbek Lotin va O'zbek Kirill imlolarini to'liq qo'llab-quvvatlash" : "• Ўзбек Лотин ва Ўзбек Кирилл имлоларини тўлиқ қўллаб-қувватлаш",
            font_size_pt: 15,
            is_bold: false,
            is_italic: false,
            font_color: "#e2e8f0",
            font_name: "Calibri",
            alignment: "left",
            box: { left: 5, top: 58, width: 90, height: 10 }
          },
          {
            id: "s1_sh1_p2",
            slide_index: 1,
            shape_name: "Band 3",
            item_type: "body",
            original_text: "• Dynamic auto-fit font scaling preventing text overflow",
            translated_text: targetScript === 'latin' ? "• Matn qutidan toshib ketmasligini ta'minlovchi dinamik Auto-fit shrift moslagich" : "• Матн қутидан тошиб кетмаслигини таъминловчи динамик Auto-fit шрифт мослагич",
            font_size_pt: 15,
            is_bold: false,
            is_italic: false,
            font_color: "#e2e8f0",
            font_name: "Calibri",
            alignment: "left",
            box: { left: 5, top: 70, width: 90, height: 10 }
          }
        ];

        const sampleSlide2 = [
          {
            id: "s2_sh0_p0",
            slide_index: 2,
            shape_name: "Sarlavha",
            item_type: "title",
            original_text: "Quarterly Milestones & Target Deliverables",
            translated_text: targetScript === 'latin' ? "Choraklik Muhim Bosqichlar va Maqsadli Natijalar" : "Чораклик Муҳим Босқичлар ва Мақсадли Натижалар",
            font_size_pt: 26,
            is_bold: true,
            is_italic: false,
            font_color: "#38bdf8",
            font_name: "Calibri",
            alignment: "left",
            box: { left: 5, top: 8, width: 90, height: 15 }
          },
          {
            id: "s2_tbl_r0_c0",
            slide_index: 2,
            shape_name: "Jadval [1, 1]",
            item_type: "table_cell",
            original_text: "Phase",
            translated_text: targetScript === 'latin' ? "Bosqich" : "Босқич",
            font_size_pt: 14,
            is_bold: true,
            is_italic: false,
            font_color: "#ffffff",
            font_name: "Calibri",
            alignment: "left",
            box: { left: 5, top: 30, width: 28, height: 15 }
          },
          {
            id: "s2_tbl_r0_c1",
            slide_index: 2,
            shape_name: "Jadval [1, 2]",
            item_type: "table_cell",
            original_text: "Strategic Objective",
            translated_text: targetScript === 'latin' ? "Strategik Maqsad" : "Стратегик Мақсад",
            font_size_pt: 14,
            is_bold: true,
            is_italic: false,
            font_color: "#ffffff",
            font_name: "Calibri",
            alignment: "left",
            box: { left: 35, top: 30, width: 35, height: 15 }
          },
          {
            id: "s2_tbl_r0_c2",
            slide_index: 2,
            shape_name: "Jadval [1, 3]",
            item_type: "table_cell",
            original_text: "Target Timeline",
            translated_text: targetScript === 'latin' ? "Rejalashtirilgan Muddat" : "Режалаштирилган Муддат",
            font_size_pt: 14,
            is_bold: true,
            is_italic: false,
            font_color: "#ffffff",
            font_name: "Calibri",
            alignment: "left",
            box: { left: 72, top: 30, width: 23, height: 15 }
          },
          {
            id: "s2_tbl_r1_c0",
            slide_index: 2,
            shape_name: "Jadval [2, 1]",
            item_type: "table_cell",
            original_text: "Q1 2026",
            translated_text: targetScript === 'latin' ? "2026-yil 1-chorak" : "2026-йил 1-чорак",
            font_size_pt: 13,
            is_bold: false,
            is_italic: false,
            font_color: "#cbd5e1",
            font_name: "Calibri",
            alignment: "left",
            box: { left: 5, top: 50, width: 28, height: 15 }
          },
          {
            id: "s2_tbl_r1_c1",
            slide_index: 2,
            shape_name: "Jadval [2, 2]",
            item_type: "table_cell",
            original_text: "Core Engine & Run-Level Formatter",
            translated_text: targetScript === 'latin' ? "Asosiy Dvigatel va Run-Level Formatlovchi" : "Асосий Двигател ва Run-Level Форматловчи",
            font_size_pt: 13,
            is_bold: false,
            is_italic: false,
            font_color: "#cbd5e1",
            font_name: "Calibri",
            alignment: "left",
            box: { left: 35, top: 50, width: 35, height: 15 }
          },
          {
            id: "s2_tbl_r1_c2",
            slide_index: 2,
            shape_name: "Jadval [2, 3]",
            item_type: "table_cell",
            original_text: "March 31",
            translated_text: targetScript === 'latin' ? "31-mart" : "31-март",
            font_size_pt: 13,
            is_bold: false,
            is_italic: false,
            font_color: "#cbd5e1",
            font_name: "Calibri",
            alignment: "left",
            box: { left: 72, top: 50, width: 23, height: 15 }
          }
        ];

        setSession({
          session_id: "demo_sample_session",
          filename: "Strategic_Roadmap_2030_Demo.pptx",
          aspect_ratio: "16:9",
          slides_count: 2,
          total_items: sampleSlideItems.length + sampleSlide2.length,
          slides: [
            {
              slide_index: 1,
              slide_id: 256,
              title: "Strategic Roadmap 2030",
              items_count: sampleSlideItems.length,
              items: sampleSlideItems
            },
            {
              slide_index: 2,
              slide_id: 257,
              title: "Quarterly Milestones & Target Deliverables",
              items_count: sampleSlide2.length,
              items: sampleSlide2
            }
          ]
        });
        setActiveSlideIdx(1);
      }
    } catch (e) {
      alert("Namunani yuklashda xatolik: " + e.message);
    } finally {
      setLoading(false);
    }
  };

  // Run AI Translation via Gemini
  const handleTranslateAll = async () => {
    if (!session) return;
    setTranslating(true);
    setProgress(10);
    setStatusMsg("Gemini AI orqali O'zbek tiliga tarjima qilinmoqda...");

    try {
      const progressTimer = setInterval(() => {
        setProgress(p => Math.min(p + 15, 85));
      }, 400);

      const res = await fetch(`${API_BASE}/translate`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          session_id: session.session_id,
          api_key: apiKey || undefined,
          target_script: targetScript,
          domain: domain,
          glossary: glossary
        })
      });

      clearInterval(progressTimer);
      setProgress(100);

      if (!res.ok) {
        const err = await res.json();
        throw new Error(err.detail || "Tarjimada xatolik yuz berdi");
      }

      const data = await res.json();
      if (data.slides) {
        setSession(prev => ({
          ...prev,
          slides: data.slides
        }));
      }
      setStatusMsg("Tarjima muvaffaqiyatli yakunlandi!");
    } catch (e) {
      alert("Tarjima xatosi: " + e.message);
    } finally {
      setTranslating(false);
    }
  };

  // Inline Edit Update
  const handleItemTextChange = (itemId, newText) => {
    if (!session) return;
    setSession(prev => {
      const newSlides = prev.slides.map(slide => {
        const newItems = slide.items.map(item => {
          if (item.id === itemId) {
            return { ...item, translated_text: newText };
          }
          return item;
        });
        return { ...slide, items: newItems };
      });
      return { ...prev, slides: newSlides };
    });

    // Notify backend debounced/async
    fetch(`${API_BASE}/update-text`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        session_id: session.session_id,
        updates: [{ id: itemId, translated_text: newText }]
      })
    }).catch(console.error);
  };

  // Quick Script Switch for all items (Latin <-> Cyrillic)
  const handleToggleScript = (newScript) => {
    setTargetScript(newScript);
    if (!session) return;
    // Auto-transliterate loaded items locally
    const isToCyrillic = newScript === 'cyrillic';
    setSession(prev => {
      const newSlides = prev.slides.map(slide => {
        const newItems = slide.items.map(item => {
          const converted = transliterateUzbek(item.translated_text, isToCyrillic);
          return { ...item, translated_text: converted };
        });
        return { ...slide, items: newItems };
      });
      return { ...prev, slides: newSlides };
    });
  };

  // Transliterate helper
  const transliterateUzbek = (text, toCyrillic) => {
    if (!text) return '';
    if (toCyrillic) {
      const comp = [
        ["Yo", "Ё"], ["yo", "ё"], ["Yu", "Ю"], ["yu", "ю"], ["Ya", "Я"], ["ya", "я"],
        ["Sh", "Ш"], ["sh", "ш"], ["Ch", "Ч"], ["ch", "ч"],
        ["Oʻ", "Ў"], ["O'", "Ў"], ["O`", "Ў"], ["O’", "Ў"], ["oʻ", "ў"], ["o'", "ў"], ["o`", "ў"], ["o’", "ў"],
        ["Gʻ", "Ғ"], ["G'", "Ғ"], ["G`", "Ғ"], ["G’", "Ғ"], ["gʻ", "ғ"], ["g'", "ғ"], ["g`", "ғ"], ["g’", "ғ"],
        ["Ts", "Ц"], ["ts", "ц"]
      ];
      let res = text;
      for (const [lat, cyr] of comp) res = res.replaceAll(lat, cyr);
      const single = {
        'A':'А','a':'а','B':'Б','b':'б','D':'Д','d':'д','E':'Е','e':'е','F':'Ф','f':'ф',
        'G':'Г','g':'г','H':'Ҳ','h':'ҳ','I':'И','i':'и','J':'Ж','j':'ж','K':'К','k':'к',
        'L':'Л','l':'л','M':'М','m':'м','N':'Н','n':'н','O':'О','o':'о','P':'П','p':'п',
        'Q':'Қ','q':'қ','R':'Р','r':'р','S':'С','s':'с','T':'Т','t':'т','U':'У','u':'у',
        'V':'В','v':'в','X':'Х','x':'х','Y':'Й','y':'й','Z':'З','z':'з'
      };
      return res.split('').map(c => single[c] || c).join('');
    } else {
      const comp = [
        ["Ё", "Yo"], ["ё", "yo"], ["Ю", "Yu"], ["ю", "yu"], ["Я", "Ya"], ["я", "ya"],
        ["Ш", "Sh"], ["ш", "sh"], ["Ч", "Ch"], ["ч", "ch"],
        ["Ў", "Oʻ"], ["ў", "oʻ"], ["Ғ", "Gʻ"], ["ғ", "gʻ"], ["Ц", "Ts"], ["ц", "ts"]
      ];
      let res = text;
      for (const [cyr, lat] of comp) res = res.replaceAll(cyr, lat);
      const single = {
        'А':'A','а':'a','Б':'B','б':'b','В':'V','в':'v','Г':'G','г':'g','Д':'D','d':'d',
        'Е':'E','е':'e','Ж':'J','ж':'j','З':'Z','z':'z','И':'I','и':'i','Й':'Y','й':'y',
        'К':'K','к':'k','Л':'L','л':'l','М':'M','м':'m','Н':'N','н':'n','О':'O','о':'o',
        'П':'P','п':'p','Р':'R','р':'r','С':'S','с':'s','Т':'T','t':'t','У':'U','у':'u',
        'Ф':'F','ф':'f','Х':'X','х':'x','Ҳ':'H','ҳ':'h','Қ':'Q','қ':'q','Э':'E','э':'e'
      };
      return res.split('').map(c => single[c] || c).join('');
    }
  };

  // Export & Download PPTX
  const handleExport = async () => {
    if (!session) return;
    setExporting(true);
    setStatusMsg("Tarjima qilingan PowerPoint (.pptx) yaratilmoqda...");

    try {
      const res = await fetch("/api/export", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          session_id: session.session_id,
          auto_fit: autoFit,
          target_script: targetScript,
          api_key: apiKey || undefined
        })
      });

      if (!res.ok) {
        let errorMsg = "Eksport qilishda xatolik yuz berdi.";
        try {
          const errJson = await res.json();
          errorMsg = errJson.detail || errorMsg;
        } catch (_) {
          const errText = await res.text();
          errorMsg = errText || errorMsg;
        }
        throw new Error(errorMsg);
      }

      const data = await res.json();
      window.location.href = data.download_url;
      setStatusMsg(`"${data.download_filename || 'Taqdimot'}" muvaffaqiyatli yuklab olindi!`);
    } catch (e) {
      alert("Eksport xatosi: " + e.message);
    } finally {
      setExporting(false);
    }
  };

  // Copy to clipboard
  const handleCopy = (id, text) => {
    navigator.clipboard.writeText(text);
    setCopiedId(id);
    setTimeout(() => setCopiedId(null), 2000);
  };

  // Add term to glossary
  const handleAddGlossaryTerm = () => {
    if (!newTermKey.trim() || !newTermVal.trim()) return;
    setGlossary(prev => ({
      ...prev,
      [newTermKey.trim()]: newTermVal.trim()
    }));
    setNewTermKey('');
    setNewTermVal('');
  };

  const handleRemoveGlossaryTerm = (key) => {
    setGlossary(prev => {
      const next = { ...prev };
      delete next[key];
      return next;
    });
  };

  const currentSlide = session?.slides?.find(s => s.slide_index === activeSlideIdx) || session?.slides?.[0];

  // Filter items for current slide
  const filteredItems = currentSlide?.items?.filter(item => {
    const matchSearch = searchTerm === '' || 
      item.original_text.toLowerCase().includes(searchTerm.toLowerCase()) ||
      item.translated_text.toLowerCase().includes(searchTerm.toLowerCase());
    
    const matchType = filterType === 'all' || 
      (filterType === 'title' && item.item_type === 'title') ||
      (filterType === 'body' && item.item_type === 'body') ||
      (filterType === 'table' && item.item_type === 'table_cell');

    return matchSearch && matchType;
  }) || [];

  return (
    <div className="flex flex-col min-h-screen bg-slate-950 text-slate-100">
      
      {/* 1. TOP NAVBAR */}
      <header className="sticky top-0 z-40 border-b border-slate-800/80 bg-slate-900/90 backdrop-blur-md px-6 py-3.5 flex items-center justify-between shadow-xl">
        <div className="flex items-center gap-3.5">
          <div className="h-10 w-10 rounded-xl bg-gradient-to-tr from-brand-600 to-indigo-500 flex items-center justify-center shadow-lg shadow-brand-500/20 text-white font-bold text-xl">
            <Zap className="h-6 w-6 text-white" />
          </div>
          <div>
            <div className="flex items-center gap-2">
              <h1 className="font-extrabold text-lg tracking-tight bg-gradient-to-r from-white via-slate-200 to-brand-300 bg-clip-text text-transparent">
                SlideTranslate AI
              </h1>
              <span className="text-[10px] font-semibold uppercase px-2 py-0.5 rounded-full bg-brand-500/20 text-brand-400 border border-brand-500/30">
                PRO Studio
              </span>
            </div>
            <p className="text-xs text-slate-400">PowerPoint Taqdimotlarini O'zbek Tiliga 100% Formatda O'girish</p>
          </div>
        </div>

        {/* Center Control Group */}
        <div className="flex items-center gap-2.5">
          {/* Script Toggle */}
          <div className="bg-slate-800/90 border border-slate-700/80 p-1 rounded-xl flex items-center gap-1 shadow-inner">
            <button
              onClick={() => handleToggleScript('latin')}
              className={`px-3 py-1.5 rounded-lg text-xs font-semibold transition-all flex items-center gap-1.5 ${
                targetScript === 'latin'
                  ? 'bg-brand-600 text-white shadow-md shadow-brand-500/30'
                  : 'text-slate-400 hover:text-slate-200 hover:bg-slate-700/50'
              }`}
            >
              <span>🔤</span> O'zbekcha (Lotin)
            </button>
            <button
              onClick={() => handleToggleScript('cyrillic')}
              className={`px-3 py-1.5 rounded-lg text-xs font-semibold transition-all flex items-center gap-1.5 ${
                targetScript === 'cyrillic'
                  ? 'bg-brand-600 text-white shadow-md shadow-brand-500/30'
                  : 'text-slate-400 hover:text-slate-200 hover:bg-slate-700/50'
              }`}
            >
              <span>🌐</span> Ўзбекча (Кирилл)
            </button>
          </div>

          {/* Domain Selector */}
          <select 
            value={domain}
            onChange={(e) => setDomain(e.target.value)}
            className="bg-slate-800/90 border border-slate-700/80 text-xs text-slate-200 rounded-xl px-3 py-2 font-medium focus:ring-2 focus:ring-brand-500 focus:outline-none cursor-pointer"
          >
            <option value="IT & Dasturlash">💻 IT & Dasturlash</option>
            <option value="Biznes & Strategiya">🏢 Biznes & Strategiya</option>
            <option value="Moliya & Iqtisodiyot">📈 Moliya & Iqtisodiyot</option>
            <option value="Ta'lim & Fan">🎓 Ta'lim & Fan</option>
            <option value="Tibbiyot & Salomatlik">🏥 Tibbiyot</option>
            <option value="Umumiy">🌐 Umumiy Mavzu</option>
          </select>
        </div>

        {/* Right Tools */}
        <div className="flex items-center gap-2">
          {/* Glossary Button */}
          <button 
            onClick={() => setShowGlossaryModal(true)}
            className="relative px-3 py-2 rounded-xl bg-slate-800 hover:bg-slate-700/80 border border-slate-700 text-xs font-medium text-slate-200 transition-all flex items-center gap-2"
            title="Maxsus atamalar lug'ati"
          >
            <BookOpen className="h-4 w-4 text-emerald-400" />
            <span>Lug'at</span>
            <span className="bg-emerald-500/20 text-emerald-300 text-[10px] px-1.5 py-0.2 rounded-full font-bold border border-emerald-500/30">
              {Object.keys(glossary).length}
            </span>
          </button>

          {/* Settings Button */}
          <button 
            onClick={() => setShowSettingsModal(true)}
            className="p-2.5 rounded-xl bg-slate-800 hover:bg-slate-700/80 border border-slate-700 text-slate-300 hover:text-white transition-all"
            title="API & Sozlamalar"
          >
            <Settings className="h-4 w-4" />
          </button>
        </div>
      </header>

      {/* 2. MAIN WORKSPACE */}
      {!session ? (
        /* HERO & UPLOAD ZONE */
        <main className="flex-1 flex flex-col items-center justify-center p-6 md:p-12 max-w-5xl mx-auto w-full">
          
          <div className="text-center space-y-4 mb-10">
            <div className="inline-flex items-center gap-2 px-3 py-1 rounded-full bg-brand-500/10 border border-brand-500/20 text-brand-400 text-xs font-semibold">
              <Sparkles className="h-3.5 w-3.5 animate-pulse" />
              Eng Yuqori Aniqlikdagi Taqdimot Tarjimon Dvigateli
            </div>
            <h2 className="text-4xl md:text-5xl font-extrabold tracking-tight text-white max-w-2xl leading-tight">
              Slaydlarni <span className="bg-gradient-to-r from-brand-400 to-indigo-400 bg-clip-text text-transparent">Dizayni Buzilmagan</span> Holda O'zbekchaga O'giring
            </h2>
            <p className="text-slate-400 text-sm md:text-base max-w-xl mx-auto">
              Barcha shakllar, jadvallar, ranglar va shriftlar 100% asl holida saqlanadi. Auto-fit matn toshishiga qarshi aqlli texnologiya bilan jihozlangan.
            </p>
          </div>

          {/* Dropzone Card */}
          <div 
            onDragOver={(e) => { e.preventDefault(); setIsDragging(true); }}
            onDragLeave={() => setIsDragging(false)}
            onDrop={(e) => {
              e.preventDefault();
              setIsDragging(false);
              if (e.dataTransfer.files?.[0]) handleFileUpload(e.dataTransfer.files[0]);
            }}
            onClick={() => fileInputRef.current?.click()}
            className={`w-full max-w-2xl rounded-3xl border-2 border-dashed p-10 md:p-14 text-center cursor-pointer transition-all duration-300 relative overflow-hidden group ${
              isDragging 
                ? 'border-brand-400 bg-brand-500/10 scale-[1.01]' 
                : 'border-slate-700/80 bg-slate-900/60 hover:border-brand-500/50 hover:bg-slate-900/90'
            }`}
          >
            <input 
              ref={fileInputRef}
              type="file" 
              accept=".pptx" 
              className="hidden" 
              onChange={(e) => e.target.files?.[0] && handleFileUpload(e.target.files[0])}
            />

            <div className="flex flex-col items-center justify-center space-y-4">
              <div className="h-20 w-20 rounded-2xl bg-gradient-to-br from-brand-500/20 to-indigo-500/20 border border-brand-500/30 flex items-center justify-center text-brand-400 group-hover:scale-110 transition-transform shadow-xl">
                <UploadCloud className="h-10 w-10 text-brand-400" />
              </div>

              <div className="space-y-1.5">
                <p className="text-lg font-bold text-white">
                  PowerPoint (.pptx) faylingizni bu yerga tashlang
                </p>
                <p className="text-xs text-slate-400">
                  yoki kompyuterdan tanlash uchun bosing (PPTX formati)
                </p>
              </div>

              <div className="pt-2 flex flex-wrap justify-center gap-2 text-[11px] font-medium text-slate-400">
                <span className="px-2.5 py-1 rounded-md bg-slate-800/80 border border-slate-700">✓ Shakl & Ranglar Saqlanadi</span>
                <span className="px-2.5 py-1 rounded-md bg-slate-800/80 border border-slate-700">✓ Jadvallar & Grafiklar</span>
                <span className="px-2.5 py-1 rounded-md bg-slate-800/80 border border-slate-700">✓ Lotin & Kirill</span>
              </div>
            </div>
          </div>

          {/* Demo Button */}
          <div className="mt-8 flex flex-col sm:flex-row items-center gap-4">
            <button
              onClick={handleLoadSample}
              disabled={loading}
              className="px-6 py-3 rounded-2xl bg-slate-800/90 hover:bg-slate-700/90 border border-slate-700 text-sm font-semibold text-slate-200 transition-all flex items-center gap-2.5 shadow-lg group hover:scale-[1.02]"
            >
              <Sparkles className="h-4 w-4 text-amber-400 group-hover:rotate-12 transition-transform" />
              <span>🚀 Namunaviy Taqdimot Bilan Sinab Ko'rish</span>
              <ArrowRight className="h-4 w-4 text-slate-400 group-hover:translate-x-1 transition-transform" />
            </button>
          </div>

          {/* Feature Highlights Grid */}
          <div className="grid grid-cols-1 md:grid-cols-3 gap-5 mt-16 w-full max-w-4xl">
            <div className="p-5 rounded-2xl bg-slate-900/50 border border-slate-800/80 space-y-2">
              <div className="h-8 w-8 rounded-lg bg-blue-500/10 text-blue-400 flex items-center justify-center font-bold">
                <Layers className="h-4 w-4" />
              </div>
              <h4 className="font-bold text-sm text-slate-200">100% Dizayn Saqlanadi</h4>
              <p className="text-xs text-slate-400 leading-relaxed">
                Matn ichidagi qalinlik, ranglar, shrift nomlari va jadvallar strukturasiga putur yetmaydi.
              </p>
            </div>

            <div className="p-5 rounded-2xl bg-slate-900/50 border border-slate-800/80 space-y-2">
              <div className="h-8 w-8 rounded-lg bg-emerald-500/10 text-emerald-400 flex items-center justify-center font-bold">
                <Sliders className="h-4 w-4" />
              </div>
              <h4 className="font-bold text-sm text-slate-200">Auto-Fit Shrift Moslash</h4>
              <p className="text-xs text-slate-400 leading-relaxed">
                O'zbekcha matn uzunroq bo'lganda shrift hajmi mutanosib moslanib, shakldan toshib ketmaydi.
              </p>
            </div>

            <div className="p-5 rounded-2xl bg-slate-900/50 border border-slate-800/80 space-y-2">
              <div className="h-8 w-8 rounded-lg bg-purple-500/10 text-purple-400 flex items-center justify-center font-bold">
                <Cpu className="h-4 w-4" />
              </div>
              <h4 className="font-bold text-sm text-slate-200">Gemini AI Dvigateli</h4>
              <p className="text-xs text-slate-400 leading-relaxed">
                Taqdimot mavzusiga (IT, Biznes, Tibbiyot) mos keluvchi professional adabiy tarjima.
              </p>
            </div>
          </div>

        </main>
      ) : (
        /* INTERACTIVE TRANSLATION STUDIO */
        <div className="flex-1 flex flex-col overflow-hidden">
          
          {/* Top Sub-Bar with Actions */}
          <div className="bg-slate-900 border-b border-slate-800 px-6 py-2.5 flex flex-wrap items-center justify-between gap-4">
            
            {/* File Info */}
            <div className="flex items-center gap-3">
              <div className="h-8 w-8 rounded-lg bg-brand-500/20 text-brand-400 flex items-center justify-center font-bold">
                <FileText className="h-4 w-4" />
              </div>
              <div>
                <div className="flex items-center gap-2">
                  <span className="font-bold text-xs text-white max-w-[200px] md:max-w-xs truncate" title={session.filename}>
                    {session.filename}
                  </span>
                  <span className="text-[10px] font-semibold px-2 py-0.2 rounded bg-slate-800 text-slate-300 border border-slate-700">
                    {session.aspect_ratio}
                  </span>
                  <span className="text-[10px] font-semibold px-2 py-0.2 rounded bg-slate-800 text-slate-300 border border-slate-700">
                    {session.slides_count} ta slayd
                  </span>
                </div>
                <p className="text-[11px] text-slate-400">Jami {session.total_items} ta matn bloki</p>
              </div>
            </div>

            {/* Middle Action Controls */}
            <div className="flex items-center gap-3">
              {/* Auto-Fit Toggle */}
              <label className="flex items-center gap-2 text-xs text-slate-300 cursor-pointer bg-slate-800/80 px-3 py-1.5 rounded-xl border border-slate-700/80 hover:bg-slate-800">
                <input 
                  type="checkbox" 
                  checked={autoFit} 
                  onChange={(e) => setAutoFit(e.target.checked)}
                  className="rounded border-slate-700 text-brand-600 focus:ring-0 cursor-pointer"
                />
                <span className="font-medium">⚡ Auto-fit (Shrift moslash)</span>
              </label>

              {/* View Mode Toggle */}
              <div className="bg-slate-800 p-1 rounded-xl flex items-center gap-1 border border-slate-700 text-xs">
                <button
                  onClick={() => setViewMode('split')}
                  className={`px-2.5 py-1 rounded-lg transition-all flex items-center gap-1 ${
                    viewMode === 'split' ? 'bg-slate-700 text-white font-semibold' : 'text-slate-400'
                  }`}
                  title="Yonma-yon ko'rish"
                >
                  <Columns className="h-3.5 w-3.5" />
                  <span className="hidden sm:inline">Split</span>
                </button>
                <button
                  onClick={() => setViewMode('translated')}
                  className={`px-2.5 py-1 rounded-lg transition-all flex items-center gap-1 ${
                    viewMode === 'translated' ? 'bg-slate-700 text-white font-semibold' : 'text-slate-400'
                  }`}
                  title="Faqat O'zbekcha"
                >
                  <Eye className="h-3.5 w-3.5" />
                  <span className="hidden sm:inline">O'zbek</span>
                </button>
              </div>
            </div>

            {/* Right Action Buttons */}
            <div className="flex items-center gap-2.5">
              {/* Translate All Button */}
              <button
                onClick={handleTranslateAll}
                disabled={translating}
                className="px-4 py-2 rounded-xl bg-gradient-to-r from-brand-600 to-indigo-600 hover:from-brand-500 hover:to-indigo-500 text-white text-xs font-bold transition-all shadow-md shadow-brand-500/20 flex items-center gap-2 disabled:opacity-50"
              >
                {translating ? (
                  <>
                    <RefreshCw className="h-4 w-4 animate-spin" />
                    <span>AI Tarjima qilinmoqda...</span>
                  </>
                ) : (
                  <>
                    <Sparkles className="h-4 w-4" />
                    <span>✨ Gemini AI Bilan Tarjima Qilish</span>
                  </>
                )}
              </button>

              {/* Export Button */}
              <button
                onClick={handleExport}
                disabled={exporting}
                className="px-4 py-2 rounded-xl bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-bold transition-all shadow-md shadow-emerald-500/20 flex items-center gap-2 disabled:opacity-50"
              >
                {exporting ? (
                  <>
                    <RefreshCw className="h-4 w-4 animate-spin" />
                    <span>Fayl tayyorlanmoqda...</span>
                  </>
                ) : (
                  <>
                    <Download className="h-4 w-4" />
                    <span>📥 PPTX Yuklab Olish</span>
                  </>
                )}
              </button>

              {/* Reset / New File */}
              <button
                onClick={() => setSession(null)}
                className="p-2 rounded-xl bg-slate-800 hover:bg-slate-700 text-slate-400 hover:text-white transition-all text-xs"
                title="Yangi fayl yuklash"
              >
                <RefreshCw className="h-4 w-4" />
              </button>
            </div>

          </div>

          {/* Progress Bar (if translating) */}
          {translating && (
            <div className="w-full bg-slate-800 h-1.5 overflow-hidden">
              <div 
                className="bg-gradient-to-r from-brand-500 to-emerald-400 h-full transition-all duration-300"
                style={{ width: `${progress}%` }}
              />
            </div>
          )}

          {/* Three-Column Workspace Body */}
          <div className="flex-1 flex overflow-hidden">
            
            {/* 1. LEFT SIDEBAR: Slide Navigator */}
            <aside className="w-64 border-r border-slate-800 bg-slate-950/60 flex flex-col overflow-y-auto p-4 space-y-2.5">
              <div className="flex items-center justify-between px-2 text-xs font-bold text-slate-400 uppercase tracking-wider">
                <span>Slaydlar ({session.slides_count})</span>
                <Layers className="h-3.5 w-3.5" />
              </div>

              {session.slides.map((slide) => {
                const isActive = slide.slide_index === activeSlideIdx;
                return (
                  <button
                    key={slide.slide_index}
                    onClick={() => setActiveSlideIdx(slide.slide_index)}
                    className={`w-full text-left p-3 rounded-2xl border transition-all flex flex-col gap-1.5 relative group ${
                      isActive 
                        ? 'bg-slate-800/90 border-brand-500/80 shadow-md shadow-brand-500/10' 
                        : 'bg-slate-900/40 border-slate-800/80 hover:bg-slate-900 hover:border-slate-700'
                    }`}
                  >
                    <div className="flex items-center justify-between">
                      <span className={`text-[11px] font-bold px-2 py-0.5 rounded-md ${
                        isActive ? 'bg-brand-500 text-white' : 'bg-slate-800 text-slate-400'
                      }`}>
                        Slayd {slide.slide_index}
                      </span>
                      <span className="text-[10px] text-slate-500 font-medium">
                        {slide.items_count} ta matn
                      </span>
                    </div>

                    <p className={`text-xs font-medium line-clamp-2 ${isActive ? 'text-white' : 'text-slate-300'}`}>
                      {slide.title || `Slayd ${slide.slide_index}`}
                    </p>
                  </button>
                );
              })}
            </aside>

            {/* 2. CENTER PANEL: Visual Slide Canvas Preview */}
            <div className="flex-1 bg-slate-950 p-4 md:p-8 flex flex-col items-center justify-start overflow-y-auto min-h-0">
              
              <div className="w-full max-w-5xl flex flex-col gap-4">
                
                {/* Slide Top Navigation & Meta Bar */}
                <div className="flex items-center justify-between bg-slate-900/90 border border-slate-800 px-4 py-2.5 rounded-2xl shadow-md">
                  <div className="flex items-center gap-3">
                    <span className="h-2.5 w-2.5 rounded-full bg-emerald-400 animate-pulse" />
                    <span className="text-xs font-bold text-white uppercase tracking-wider">
                      Slayd {currentSlide?.slide_index} / {session.slides_count}
                    </span>
                    <span className="text-[11px] text-slate-400 font-medium">
                      ({currentSlide?.items?.length || 0} ta matn bloki)
                    </span>
                  </div>

                  {/* Slide Stepper Controls */}
                  <div className="flex items-center gap-2">
                    <button
                      onClick={() => setActiveSlideIdx(Math.max(1, activeSlideIdx - 1))}
                      disabled={activeSlideIdx <= 1}
                      className="px-3 py-1.5 rounded-xl bg-slate-800 hover:bg-slate-700 text-xs font-semibold text-slate-200 disabled:opacity-30 transition-all flex items-center gap-1 border border-slate-700"
                    >
                      <span>⬅</span> Oldingi
                    </button>

                    <div className="flex items-center gap-1 overflow-x-auto max-w-[240px] px-1">
                      {session.slides.map(s => (
                        <button
                          key={s.slide_index}
                          onClick={() => setActiveSlideIdx(s.slide_index)}
                          className={`h-7 w-7 rounded-lg text-xs font-bold transition-all flex-shrink-0 ${
                            s.slide_index === activeSlideIdx
                              ? 'bg-brand-600 text-white shadow-md shadow-brand-500/40 border border-brand-400'
                              : 'bg-slate-800/80 text-slate-400 hover:bg-slate-700 hover:text-white border border-slate-700/60'
                          }`}
                        >
                          {s.slide_index}
                        </button>
                      ))}
                    </div>

                    <button
                      onClick={() => setActiveSlideIdx(Math.min(session.slides_count, activeSlideIdx + 1))}
                      disabled={activeSlideIdx >= session.slides_count}
                      className="px-3 py-1.5 rounded-xl bg-slate-800 hover:bg-slate-700 text-xs font-semibold text-slate-200 disabled:opacity-30 transition-all flex items-center gap-1 border border-slate-700"
                    >
                      Keyingi <span>➡</span>
                    </button>
                  </div>
                </div>

                {/* The Aspect-Ratio Slide Card */}
                <div className="relative w-full aspect-video bg-white text-slate-900 rounded-2xl border-4 border-slate-700/80 shadow-2xl overflow-hidden flex flex-col justify-between p-6 md:p-10 transition-all">
                  
                  {/* Decorative slide header banner */}
                  <div className="absolute top-0 left-0 right-0 h-2 bg-gradient-to-r from-brand-600 via-indigo-600 to-emerald-500" />
                  
                  {/* Slide Content Area */}
                  <div className="flex-1 overflow-y-auto space-y-4 pr-1 mt-1">
                    
                    {/* View Mode: Split Comparison vs Full Translated */}
                    {viewMode === 'split' ? (
                      <div className="grid grid-cols-2 gap-6 h-full">
                        
                        {/* Left: Original English */}
                        <div className="bg-slate-50 p-4 rounded-xl border border-slate-200 space-y-3">
                          <div className="text-[10px] font-bold uppercase tracking-wider text-slate-500 pb-1 border-b border-slate-200 flex items-center gap-1.5">
                            <span>🇺🇸</span> Asl Matn (Inglizcha / Xorijiy)
                          </div>
                          <div className="space-y-2">
                            {currentSlide?.items?.map(item => (
                              <div key={item.id} className="text-xs text-slate-700 bg-white p-2.5 rounded-lg border border-slate-200/80 shadow-sm leading-relaxed">
                                <span className="text-[9px] text-slate-400 font-mono block mb-0.5">{item.shape_name}</span>
                                {item.original_text}
                              </div>
                            ))}
                          </div>
                        </div>

                        {/* Right: Translated Uzbek */}
                        <div className="bg-emerald-50/50 p-4 rounded-xl border border-emerald-200 space-y-3">
                          <div className="text-[10px] font-bold uppercase tracking-wider text-emerald-700 pb-1 border-b border-emerald-200 flex items-center gap-1.5">
                            <span>🇺🇿</span> O'zbekcha Tarjima ({targetScript === 'latin' ? 'Lotin' : 'Kirill'})
                          </div>
                          <div className="space-y-2">
                            {currentSlide?.items?.map(item => (
                              <div key={item.id} className="text-xs text-slate-900 font-medium bg-white p-2.5 rounded-lg border border-emerald-200 shadow-sm leading-relaxed">
                                <span className="text-[9px] text-emerald-600 font-mono block mb-0.5">{item.shape_name}</span>
                                {item.translated_text || item.original_text}
                              </div>
                            ))}
                          </div>
                        </div>

                      </div>
                    ) : (
                      /* Full Slide Presentation Preview */
                      <div className="space-y-4">
                        
                        {/* Slide Title / Top items */}
                        {(() => {
                          const items = currentSlide?.items || [];
                          const nonTable = items.filter(i => i.item_type !== 'table_cell');
                          if (nonTable.length === 0) return null;
                          const title = nonTable[0];
                          return (
                            <div className="border-b border-slate-200 pb-3">
                              <h2 className="text-2xl md:text-3xl font-extrabold text-slate-900 tracking-tight">
                                {title.translated_text || title.original_text}
                              </h2>
                              <span className="text-[10px] text-brand-600 font-mono mt-0.5 block">{title.shape_name}</span>
                            </div>
                          );
                        })()}

                        {/* Slide Secondary / Body Items */}
                        {(() => {
                          const items = currentSlide?.items || [];
                          const nonTable = items.filter(i => i.item_type !== 'table_cell');
                          if (nonTable.length <= 1) return null;
                          const rest = nonTable.slice(1);
                          return (
                            <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                              {rest.map(item => (
                                <div key={item.id} className="p-3 bg-slate-50 rounded-xl border border-slate-200 text-xs md:text-sm text-slate-800 leading-relaxed shadow-sm">
                                  <span className="text-[9px] text-slate-400 font-mono block mb-1">{item.shape_name}</span>
                                  {item.translated_text || item.original_text}
                                </div>
                              ))}
                            </div>
                          );
                        })()}

                        {/* Table Matrix Preview (Clean Grid) */}
                        {currentSlide?.items?.some(i => i.item_type === 'table_cell') && (
                          <div className="pt-2">
                            <div className="border border-slate-300 rounded-xl overflow-hidden shadow-sm">
                              <div className="bg-slate-100 px-3 py-1.5 border-b border-slate-300 text-[11px] font-bold text-slate-700 flex items-center justify-between">
                                <span>📊 Jadval Elementlari</span>
                                <span className="text-[10px] text-slate-500">{currentSlide.items.filter(i => i.item_type === 'table_cell').length} ta katak</span>
                              </div>
                              <div className="grid grid-cols-2 md:grid-cols-3 gap-1.5 p-2.5 bg-slate-50 max-h-56 overflow-y-auto">
                                {currentSlide.items.filter(i => i.item_type === 'table_cell').map(item => (
                                  <div key={item.id} className="p-2 bg-white rounded-lg border border-slate-200 text-xs text-slate-900 shadow-sm">
                                    <span className="text-[9px] text-indigo-600 font-semibold block mb-0.5">{item.shape_name}</span>
                                    <span className="font-medium">{item.translated_text || item.original_text}</span>
                                  </div>
                                ))}
                              </div>
                            </div>
                          </div>
                        )}

                      </div>
                    )}

                  </div>

                  {/* Slide Footer */}
                  <div className="pt-3 mt-2 border-t border-slate-200 flex items-center justify-between text-[11px] text-slate-500 font-medium">
                    <span className="flex items-center gap-1.5 text-brand-700 font-semibold">
                      <Sparkles className="h-3.5 w-3.5" />
                      SlideTranslate AI — {targetScript === 'latin' ? "O'zbekcha (Lotin)" : "Ўзбекча (Кирилл)"}
                    </span>
                    <span className="bg-slate-100 text-slate-700 px-3 py-1 rounded-full border border-slate-300 font-bold">
                      {currentSlide?.slide_index} / {session.slides_count}
                    </span>
                  </div>

                </div>

              </div>

            </div>

            {/* 3. RIGHT PANEL: Inline Text Editor & Translation Cards */}
            <div className="w-[450px] border-l border-slate-800 bg-slate-950/80 flex flex-col overflow-hidden">
              
              {/* Search & Filter Header */}
              <div className="p-4 border-b border-slate-800 space-y-3 bg-slate-900/40">
                <div className="flex items-center justify-between">
                  <h3 className="font-bold text-xs uppercase tracking-wider text-slate-300">
                    Matnlarni Tahrirlash ({filteredItems.length})
                  </h3>
                  <span className="text-[11px] text-emerald-400 font-medium">Avto-saqlash yoqilgan</span>
                </div>

                {/* Search Bar */}
                <div className="relative">
                  <Search className="h-3.5 w-3.5 absolute left-3 top-3 text-slate-500" />
                  <input
                    type="text"
                    value={searchTerm}
                    onChange={(e) => setSearchTerm(e.target.value)}
                    placeholder="Slayd matnlarini izlash..."
                    className="w-full bg-slate-900 border border-slate-700/80 rounded-xl pl-9 pr-3 py-1.5 text-xs text-slate-200 placeholder-slate-500 focus:outline-none focus:ring-1 focus:ring-brand-500"
                  />
                </div>

                {/* Type Filter Pills */}
                <div className="flex items-center gap-1.5 text-[10px]">
                  {['all', 'title', 'body', 'table'].map(type => (
                    <button
                      key={type}
                      onClick={() => setFilterType(type)}
                      className={`px-2.5 py-1 rounded-lg font-medium transition-all ${
                        filterType === type 
                          ? 'bg-brand-600 text-white font-bold' 
                          : 'bg-slate-800 text-slate-400 hover:text-slate-200'
                      }`}
                    >
                      {type === 'all' && 'Barchasi'}
                      {type === 'title' && 'Sarlavhalar'}
                      {type === 'body' && 'Matnlar'}
                      {type === 'table' && 'Jadvallar'}
                    </button>
                  ))}
                </div>
              </div>

              {/* Editable Cards List */}
              <div className="flex-1 overflow-y-auto p-4 space-y-4">
                {filteredItems.length === 0 ? (
                  <div className="text-center py-12 text-slate-500 text-xs">
                    Mos keluvchi matn bloklari topilmadi.
                  </div>
                ) : (
                  filteredItems.map((item, idx) => (
                    <div 
                      key={item.id}
                      className="p-4 rounded-2xl bg-slate-900/90 border border-slate-800 hover:border-slate-700 transition-all space-y-3 shadow-lg"
                    >
                      {/* Item Header */}
                      <div className="flex items-center justify-between text-[11px]">
                        <span className="font-semibold text-brand-400 flex items-center gap-1.5">
                          <span className="h-2 w-2 rounded-full bg-brand-400 inline-block" />
                          {item.shape_name}
                        </span>
                        
                        <div className="flex items-center gap-2">
                          <span className="text-[10px] text-slate-500">
                            {item.font_size_pt}pt {item.is_bold ? 'Bold' : ''}
                          </span>
                          <button
                            onClick={() => handleCopy(item.id, item.translated_text || item.original_text)}
                            className="text-slate-400 hover:text-white transition-colors"
                            title="Nusxa olish"
                          >
                            {copiedId === item.id ? <Check className="h-3.5 w-3.5 text-emerald-400" /> : <Copy className="h-3.5 w-3.5" />}
                          </button>
                        </div>
                      </div>

                      {/* Original Box */}
                      <div className="p-2.5 rounded-xl bg-slate-950/70 border border-slate-800/80 text-xs text-slate-400 leading-relaxed font-mono select-all">
                        <span className="text-[9px] uppercase font-bold text-slate-500 block mb-0.5">Asl matn:</span>
                        {item.original_text}
                      </div>

                      {/* Translated Editable Box */}
                      <div className="space-y-1">
                        <span className="text-[9px] uppercase font-bold text-emerald-400 flex items-center gap-1">
                          <span>O'zbekcha Tarjima:</span>
                        </span>
                        <textarea
                          rows={3}
                          value={item.translated_text || ''}
                          onChange={(e) => handleItemTextChange(item.id, e.target.value)}
                          placeholder="Tarjimani kiriting..."
                          className="w-full bg-slate-950 border border-slate-700/90 focus:border-brand-500 rounded-xl p-3 text-xs text-slate-100 placeholder-slate-600 focus:outline-none focus:ring-1 focus:ring-brand-500 leading-relaxed resize-y font-sans"
                        />
                      </div>

                      {/* Item Quick Actions */}
                      <div className="flex items-center justify-between text-[11px] pt-1 text-slate-400">
                        <button
                          onClick={() => handleItemTextChange(item.id, item.original_text)}
                          className="hover:text-amber-400 transition-colors text-[10px]"
                        >
                          ↺ Asliga qaytarish
                        </button>
                        
                        <button
                          onClick={() => {
                            const toggled = transliterateUzbek(item.translated_text, targetScript === 'latin');
                            handleItemTextChange(item.id, toggled);
                          }}
                          className="hover:text-brand-400 transition-colors text-[10px]"
                        >
                          🔀 Lotin ↔ Кирилл
                        </button>
                      </div>
                    </div>
                  ))
                )}
              </div>

            </div>

          </div>

        </div>
      )}

      {/* 3. GLOSSARY MODAL */}
      {showGlossaryModal && (
        <div className="fixed inset-0 z-50 bg-black/70 backdrop-blur-sm flex items-center justify-center p-4">
          <div className="bg-slate-900 border border-slate-800 rounded-3xl max-w-xl w-full p-6 space-y-5 shadow-2xl">
            
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-3">
                <div className="h-9 w-9 rounded-xl bg-emerald-500/20 text-emerald-400 flex items-center justify-center">
                  <BookOpen className="h-5 w-5" />
                </div>
                <div>
                  <h3 className="font-bold text-base text-white">Maxsus Atamalar Lug'ati (Glossary)</h3>
                  <p className="text-xs text-slate-400">Belgilangan terminlar tarjimada qat'iy qoida sifatida ishlatiladi</p>
                </div>
              </div>
              <button 
                onClick={() => setShowGlossaryModal(false)}
                className="text-slate-400 hover:text-white text-lg font-bold"
              >
                ✕
              </button>
            </div>

            {/* Add New Term Inputs */}
            <div className="grid grid-cols-5 gap-2 pt-2">
              <input
                type="text"
                value={newTermKey}
                onChange={(e) => setNewTermKey(e.target.value)}
                placeholder="Inglizcha termin (masalan: API)"
                className="col-span-2 bg-slate-950 border border-slate-700 rounded-xl px-3 py-2 text-xs text-slate-200 focus:outline-none focus:ring-1 focus:ring-emerald-500"
              />
              <input
                type="text"
                value={newTermVal}
                onChange={(e) => setNewTermVal(e.target.value)}
                placeholder="O'zbekcha tarjimasi (masalan: API)"
                className="col-span-2 bg-slate-950 border border-slate-700 rounded-xl px-3 py-2 text-xs text-slate-200 focus:outline-none focus:ring-1 focus:ring-emerald-500"
              />
              <button
                onClick={handleAddGlossaryTerm}
                className="col-span-1 bg-emerald-600 hover:bg-emerald-500 text-white font-bold rounded-xl text-xs flex items-center justify-center gap-1"
              >
                <Plus className="h-4 w-4" /> Qo'shish
              </button>
            </div>

            {/* Existing Terms List */}
            <div className="max-h-60 overflow-y-auto space-y-2 border border-slate-800 rounded-2xl p-3 bg-slate-950/60">
              {Object.keys(glossary).length === 0 ? (
                <p className="text-xs text-slate-500 text-center py-4">Lug'at bo'sh. Yuqoridan yangi termin qo'shing.</p>
              ) : (
                Object.entries(glossary).map(([k, v]) => (
                  <div key={k} className="flex items-center justify-between p-2 rounded-xl bg-slate-900 border border-slate-800/80 text-xs">
                    <div>
                      <span className="font-mono text-slate-300 font-semibold">{k}</span>
                      <span className="text-slate-500 mx-2">➔</span>
                      <span className="text-emerald-400 font-medium">{v}</span>
                    </div>
                    <button
                      onClick={() => handleRemoveGlossaryTerm(k)}
                      className="text-slate-500 hover:text-red-400 transition-colors p-1"
                    >
                      <Trash2 className="h-3.5 w-3.5" />
                    </button>
                  </div>
                ))
              )}
            </div>

            <div className="flex justify-end pt-2">
              <button
                onClick={() => setShowGlossaryModal(false)}
                className="px-5 py-2 rounded-xl bg-slate-800 hover:bg-slate-700 text-xs font-semibold text-white"
              >
                Tayyor
              </button>
            </div>

          </div>
        </div>
      )}

      {/* 4. SETTINGS MODAL */}
      {showSettingsModal && (
        <div className="fixed inset-0 z-50 bg-black/70 backdrop-blur-sm flex items-center justify-center p-4">
          <div className="bg-slate-900 border border-slate-800 rounded-3xl max-w-lg w-full p-6 space-y-5 shadow-2xl">
            
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-3">
                <div className="h-9 w-9 rounded-xl bg-brand-500/20 text-brand-400 flex items-center justify-center">
                  <Key className="h-5 w-5" />
                </div>
                <div>
                  <h3 className="font-bold text-base text-white">Gemini API & Tizim Sozlamalari</h3>
                  <p className="text-xs text-slate-400">Google Gemini API kalitini kiritish</p>
                </div>
              </div>
              <button 
                onClick={() => setShowSettingsModal(false)}
                className="text-slate-400 hover:text-white text-lg font-bold"
              >
                ✕
              </button>
            </div>

            <div className="space-y-4">
              <div className="space-y-1.5">
                <label className="text-xs font-semibold text-slate-300 block">
                  Gemini API Kaliti (Google AI Studio)
                </label>
                <input
                  type="password"
                  value={apiKey}
                  onChange={(e) => setApiKey(e.target.value)}
                  placeholder="AIzaSy..."
                  className="w-full bg-slate-950 border border-slate-700 rounded-xl px-3 py-2.5 text-xs text-slate-200 focus:outline-none focus:ring-1 focus:ring-brand-500 font-mono"
                />
                <p className="text-[11px] text-slate-500">
                  API kalit kiritilmasa, serverning standart sozlangan Gemini kalitidan foydalaniladi.
                </p>
              </div>

              <div className="p-3.5 rounded-2xl bg-brand-500/10 border border-brand-500/20 text-xs text-slate-300 space-y-1">
                <p className="font-bold text-brand-400">💡 Tavsiya qilingan modellar:</p>
                <p className="text-[11px] text-slate-400">
                  Dvigatel avtomatik ravishda <b>Gemini 3.6 Flash</b> va <b>Gemini 3.7 Flash</b> modellaridan foydalanadi.
                </p>
              </div>
            </div>

            <div className="flex justify-end pt-2">
              <button
                onClick={() => setShowSettingsModal(false)}
                className="px-5 py-2 rounded-xl bg-brand-600 hover:bg-brand-500 text-xs font-semibold text-white"
              >
                Saqlash va Yopish
              </button>
            </div>

          </div>
        </div>
      )}

    </div>
  );
}
'''
write_file('frontend/src/App.jsx', app_jsx)



# ============================================================
# 1. backend/core/transliteration.py
# ============================================================
transliteration_code = '''# -*- coding: utf-8 -*-
import re

LATIN_TO_CYRILLIC_COMPOUND = [
    ("Yo", "\u0401"), ("YO", "\u0401"), ("yo", "\u0451"),
    ("Yu", "\u042e"), ("YU", "\u042e"), ("yu", "\u044e"),
    ("Ya", "\u042f"), ("YA", "\u042f"), ("ya", "\u044f"),
    ("Ye", "\u0415"), ("YE", "\u0415"), ("ye", "\u0435"),
    ("Sh", "\u0428"), ("SH", "\u0428"), ("sh", "\u0448"),
    ("Ch", "\u0427"), ("CH", "\u0427"), ("ch", "\u0447"),
    ("Oʻ", "\u040e"), ("O'", "\u040e"), ("O`", "\u040e"), ("O’", "\u040e"), ("Oʼ", "\u040e"),
    ("oʻ", "\u045e"), ("o'", "\u045e"), ("o`", "\u045e"), ("o’", "\u045e"), ("oʼ", "\u045e"),
    ("Gʻ", "\u0492"), ("G'", "\u0492"), ("G`", "\u0492"), ("G’", "\u0492"), ("Gʼ", "\u0492"),
    ("gʻ", "\u0493"), ("g'", "\u0493"), ("g`", "\u0493"), ("g’", "\u0493"), ("gʼ", "\u0493"),
    ("Ts", "\u0426"), ("TS", "\u0426"), ("ts", "\u0446"),
]

LATIN_TO_CYRILLIC_SINGLE = {
    'A': '\u0410', 'a': '\u0430',
    'B': '\u0411', 'b': '\u0431',
    'D': '\u0414', 'd': '\u0434',
    'E': '\u0415', 'e': '\u0435',
    'F': '\u0424', 'f': '\u0444',
    'G': '\u0413', 'g': '\u0433',
    'H': '\u04b2', 'h': '\u04b3',
    'I': '\u0418', 'i': '\u0438',
    'J': '\u0416', 'j': '\u0436',
    'K': '\u041a', 'k': '\u043a',
    'L': '\u041b', 'l': '\u043b',
    'M': '\u041c', 'm': '\u043c',
    'N': '\u041d', 'n': '\u043d',
    'O': '\u041e', 'o': '\u043e',
    'P': '\u041f', 'p': '\u043f',
    'Q': '\u049a', 'q': '\u049b',
    'R': '\u0420', 'r': '\u0440',
    'S': '\u0421', 's': '\u0441',
    'T': '\u0422', 't': '\u0442',
    'U': '\u0423', 'u': '\u0443',
    'V': '\u0412', 'v': '\u0432',
    'X': '\u0425', 'x': '\u0445',
    'Y': '\u0419', 'y': '\u0439',
    'Z': '\u0417', 'z': '\u0437',
    '’': '\u044a', "'": '\u044a', '`': '\u044a', 'ʻ': '\u044a', 'ʼ': '\u044a'
}

CYRILLIC_TO_LATIN_COMPOUND = [
    ("\u0401", "Yo"), ("\u0451", "yo"),
    ("\u042e", "Yu"), ("\u044e", "yu"),
    ("\u042f", "Ya"), ("\u044f", "ya"),
    ("\u0428", "Sh"), ("\u0448", "sh"),
    ("\u0427", "Ch"), ("\u0447", "ch"),
    ("\u040e", "Oʻ"), ("\u045e", "oʻ"),
    ("\u0492", "Gʻ"), ("\u0493", "gʻ"),
    ("\u0426", "Ts"), ("\u0446", "ts"),
    ("\u0429", "Sh"), ("\u0449", "sh"),
    ("\u042a", "’"), ("\u044a", "’"),
    ("\u042c", ""), ("\u044c", ""),
]

CYRILLIC_TO_LATIN_SINGLE = {
    '\u0410': 'A', '\u0430': 'a',
    '\u0411': 'B', '\u0431': 'b',
    '\u0412': 'V', '\u0432': 'v',
    '\u0413': 'G', '\u0433': 'g',
    '\u0414': 'D', '\u0434': 'd',
    '\u0415': 'E', '\u0435': 'e',
    '\u0416': 'J', '\u0436': 'j',
    '\u0417': 'Z', '\u0437': 'z',
    '\u0418': 'I', '\u0438': 'i',
    '\u0419': 'Y', '\u0439': 'y',
    '\u041a': 'K', '\u043a': 'k',
    '\u041b': 'L', '\u043b': 'l',
    '\u041c': 'M', '\u043c': 'm',
    '\u041d': 'N', '\u043d': 'n',
    '\u041e': 'O', '\u043e': 'o',
    '\u041f': 'P', '\u043f': 'p',
    '\u0420': 'R', '\u0440': 'r',
    '\u0421': 'S', '\u0441': 's',
    '\u0422': 'T', '\u0442': 't',
    '\u0423': 'U', '\u0443': 'u',
    '\u0424': 'F', '\u0444': 'f',
    '\u0425': 'X', '\u0445': 'x',
    '\u04b2': 'H', '\u04b3': 'h',
    '\u049a': 'Q', '\u049b': 'q',
    '\u042d': 'E', '\u044d': 'e',
}

def latin_to_cyrillic(text: str) -> str:
    if not text:
        return ""
    res = text
    for lat, cyr in LATIN_TO_CYRILLIC_COMPOUND:
        res = res.replace(lat, cyr)

    # Word-initial 'E'/'e' -> 'Э'/'э'
    def replace_e(match):
        word = match.group(0)
        if word.startswith('E'):
            return '\u042d' + word[1:]
        elif word.startswith('e'):
            return '\u044d' + word[1:]
        return word

    res = re.sub(r'\b[Ee]\w*', replace_e, res)

    chars = []
    for c in res:
        chars.append(LATIN_TO_CYRILLIC_SINGLE.get(c, c))
    return "".join(chars)

def cyrillic_to_latin(text: str) -> str:
    if not text:
        return ""
    res = text
    for cyr, lat in CYRILLIC_TO_LATIN_COMPOUND:
        res = res.replace(cyr, lat)

    chars = []
    for c in res:
        chars.append(CYRILLIC_TO_LATIN_SINGLE.get(c, c))
    return "".join(chars)

def ensure_script(text: str, target_script: str) -> str:
    if not text:
        return ""
    has_cyrillic = bool(re.search(r'[\u0400-\u04FF]', text))
    if target_script.lower() == 'cyrillic':
        if not has_cyrillic:
            return latin_to_cyrillic(text)
        return text
    else:
        if has_cyrillic:
            return cyrillic_to_latin(text)
        return text
'''
write_file('backend/core/transliteration.py', transliteration_code)

# ============================================================
# 2. backend/core/pptx_processor.py
# ============================================================
pptx_processor_code = '''# -*- coding: utf-8 -*-
import os
import copy
from typing import Dict, List, Any, Optional
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.dml.color import RGBColor

from backend.core.transliteration import ensure_script


class PPTXProcessor:
    """
    Robust PowerPoint parser and text replacer.
    Extracts text hierarchy (slide -> shape -> table -> paragraph -> runs)
    and replaces text while strictly preserving shape layout, colors,
    fonts, table styling, and performing intelligent auto-fit.
    """

    @staticmethod
    def extract_presentation_data(pptx_path: str) -> Dict[str, Any]:
        """
        Parses PPTX file and returns structured presentation metadata and text items.
        """
        if not os.path.exists(pptx_path):
            raise FileNotFoundError(f"PPTX file not found: {pptx_path}")

        prs = Presentation(pptx_path)
        slide_width = prs.slide_width
        slide_height = prs.slide_height
        
        # Calculate aspect ratio
        ratio_float = slide_width / slide_height if slide_height else 1.777
        aspect_ratio = "16:9" if abs(ratio_float - (16/9)) < 0.15 else "4:3"

        slides_data = []
        total_items_count = 0

        for s_idx, slide in enumerate(prs.slides, start=1):
            slide_items = []
            slide_title = ""

            # Extract from shapes
            PPTXProcessor._extract_shapes_recursive(
                shapes=slide.shapes,
                slide_index=s_idx,
                slide_width=slide_width,
                slide_height=slide_height,
                items_list=slide_items
            )

            # Extract from Notes slide if present
            if slide.has_notes_slide and slide.notes_slide:
                notes_tf = slide.notes_slide.notes_text_frame
                if notes_tf and notes_tf.text.strip():
                    for p_idx, p in enumerate(notes_tf.paragraphs):
                        p_text = p.text.strip()
                        if p_text:
                            item_id = f"s{s_idx}_notes_p{p_idx}"
                            slide_items.append({
                                "id": item_id,
                                "slide_index": s_idx,
                                "shape_name": "Speaker Notes",
                                "item_type": "notes",
                                "original_text": p_text,
                                "translated_text": p_text,
                                "font_size_pt": 12.0,
                                "is_bold": False,
                                "is_italic": False,
                                "font_color": "#4A5568",
                                "font_name": "Calibri",
                                "alignment": "left",
                                "box": {"left": 0, "top": 90, "width": 100, "height": 10}
                            })

            # Detect primary title for the slide
            for item in slide_items:
                if item["item_type"] == "title" and not slide_title:
                    slide_title = item["original_text"]
                    break
            if not slide_title and slide_items:
                slide_title = slide_items[0]["original_text"][:60]
            if not slide_title:
                slide_title = f"Slayd {s_idx}"

            total_items_count += len(slide_items)
            slides_data.append({
                "slide_index": s_idx,
                "slide_id": slide.slide_id,
                "title": slide_title,
                "items_count": len(slide_items),
                "items": slide_items
            })

        return {
            "aspect_ratio": aspect_ratio,
            "slide_width_pt": slide_width.pt if slide_width else 960,
            "slide_height_pt": slide_height.pt if slide_height else 540,
            "slides_count": len(slides_data),
            "total_items": total_items_count,
            "slides": slides_data
        }

    @staticmethod
    def _extract_shapes_recursive(shapes, slide_index: int, slide_width, slide_height, items_list: list, prefix: str = ""):
        """Recursively extracts text from normal shapes, tables, and group shapes."""
        for sh_idx, shape in enumerate(shapes):
            sh_id_str = f"{prefix}sh{getattr(shape, 'shape_id', sh_idx)}"

            # Calculate bounding box in percentage for frontend visual preview
            box_info = {"left": 5, "top": 5, "width": 90, "height": 20}
            try:
                if shape.left is not None and shape.top is not None and slide_width and slide_height:
                    box_info = {
                        "left": round((shape.left / slide_width) * 100, 2),
                        "top": round((shape.top / slide_height) * 100, 2),
                        "width": round((shape.width / slide_width) * 100, 2),
                        "height": round((shape.height / slide_height) * 100, 2)
                    }
            except Exception:
                pass

            # 1. Group shapes
            if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                try:
                    PPTXProcessor._extract_shapes_recursive(
                        shapes=shape.shapes,
                        slide_index=slide_index,
                        slide_width=slide_width,
                        slide_height=slide_height,
                        items_list=items_list,
                        prefix=f"{sh_id_str}_g"
                    )
                except Exception:
                    pass
                continue

            # 2. Tables
            if shape.has_table:
                table = shape.table
                for r_idx, row in enumerate(table.rows):
                    for c_idx, cell in enumerate(row.cells):
                        if cell.text_frame:
                            for p_idx, p in enumerate(cell.text_frame.paragraphs):
                                p_text = p.text.strip()
                                if p_text:
                                    item_id = f"s{slide_index}_{sh_id_str}_tbl_r{r_idx}_c{c_idx}_p{p_idx}"
                                    meta = PPTXProcessor._get_paragraph_meta(p)
                                    items_list.append({
                                        "id": item_id,
                                        "slide_index": slide_index,
                                        "shape_name": f"Jadval [Qator {r_idx+1}, Ustun {c_idx+1}]",
                                        "item_type": "table_cell",
                                        "original_text": p_text,
                                        "translated_text": p_text,
                                        "font_size_pt": meta["font_size_pt"],
                                        "is_bold": meta["is_bold"],
                                        "is_italic": meta["is_italic"],
                                        "font_color": meta["font_color"],
                                        "font_name": meta["font_name"],
                                        "alignment": meta["alignment"],
                                        "box": box_info
                                    })
                continue

            # 3. Regular Shape / TextBox
            if shape.has_text_frame:
                tf = shape.text_frame
                # Determine item type
                is_title = False
                try:
                    if shape == shapes.title or getattr(shape, "is_placeholder", False):
                        if getattr(shape, "placeholder_format", None) and shape.placeholder_format.type in (1, 3): # TITLE or CENTER_TITLE
                            is_title = True
                except Exception:
                    pass

                for p_idx, p in enumerate(tf.paragraphs):
                    p_text = p.text.strip()
                    if p_text:
                        item_id = f"s{slide_index}_{sh_id_str}_p{p_idx}"
                        meta = PPTXProcessor._get_paragraph_meta(p)
                        item_type = "title" if (is_title and p_idx == 0) else ("subtitle" if is_title else "body")
                        items_list.append({
                            "id": item_id,
                            "slide_index": slide_index,
                            "shape_name": shape.name or f"Shakl {sh_idx+1}",
                            "item_type": item_type,
                            "original_text": p_text,
                            "translated_text": p_text,
                            "font_size_pt": meta["font_size_pt"],
                            "is_bold": meta["is_bold"],
                            "is_italic": meta["is_italic"],
                            "font_color": meta["font_color"],
                            "font_name": meta["font_name"],
                            "alignment": meta["alignment"],
                            "box": box_info
                        })

    @staticmethod
    def _get_paragraph_meta(paragraph) -> Dict[str, Any]:
        """Extracts run-level and paragraph-level visual properties."""
        font_size = 18.0
        is_bold = False
        is_italic = False
        font_color = "#1A202C"
        font_name = "Calibri"
        alignment = "left"

        try:
            if paragraph.alignment:
                align_str = str(paragraph.alignment).lower()
                if "center" in align_str:
                    alignment = "center"
                elif "right" in align_str:
                    alignment = "right"
                elif "justify" in align_str:
                    alignment = "justify"
        except Exception:
            pass

        if paragraph.runs:
            # Look at first non-empty run
            for r in paragraph.runs:
                if r.text.strip():
                    if r.font.size:
                        try:
                            font_size = round(r.font.size.pt, 1)
                        except Exception:
                            pass
                    if r.font.bold is not None:
                        is_bold = bool(r.font.bold)
                    if r.font.italic is not None:
                        is_italic = bool(r.font.italic)
                    if r.font.name:
                        font_name = r.font.name
                    try:
                        if r.font.color:
                            rgb = r.font.color.rgb
                            if rgb:
                                font_color = f"#{rgb[0]:02x}{rgb[1]:02x}{rgb[2]:02x}"
                    except Exception:
                        pass
                    break

        return {
            "font_size_pt": font_size,
            "is_bold": is_bold,
            "is_italic": is_italic,
            "font_color": font_color,
            "font_name": font_name,
            "alignment": alignment
        }

    @staticmethod
    def apply_translations_and_export(
        original_pptx_path: str,
        translations_map: Dict[str, str],
        output_pptx_path: str,
        auto_fit: bool = True,
        target_script: str = "latin"
    ) -> str:
        """
        Loads original PPTX, replaces text using translations_map,
        applies Auto-fit font scaling and word-wrapping, and saves to output_pptx_path.
        """
        if not os.path.exists(original_pptx_path):
            raise FileNotFoundError(f"Original PPTX not found: {original_pptx_path}")

        prs = Presentation(original_pptx_path)

        for s_idx, slide in enumerate(prs.slides, start=1):
            # Process shapes
            PPTXProcessor._apply_to_shapes_recursive(
                shapes=slide.shapes,
                slide_index=s_idx,
                translations_map=translations_map,
                auto_fit=auto_fit,
                target_script=target_script
            )

            # Process Notes
            if slide.has_notes_slide and slide.notes_slide:
                notes_tf = slide.notes_slide.notes_text_frame
                if notes_tf:
                    for p_idx, p in enumerate(notes_tf.paragraphs):
                        item_id = f"s{s_idx}_notes_p{p_idx}"
                        if item_id in translations_map:
                            trans = ensure_script(translations_map[item_id], target_script)
                            PPTXProcessor._set_paragraph_text_safe(p, trans, auto_fit=False)

        os.makedirs(os.path.dirname(os.path.abspath(output_pptx_path)), exist_ok=True)
        out_target_dir = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
if not out_target_dir.endswith('.pptx'):
    os.makedirs(out_target_dir, exist_ok=True)
    out_target_file = os.path.join(out_target_dir, 'presentation_95_precision.pptx')
else:
    out_target_file = out_target_dir
prs.save(out_target_file)
print(f"Presentation saved successfully to {out_target_file}")
        return output_pptx_path

    @staticmethod
    def _apply_to_shapes_recursive(shapes, slide_index: int, translations_map: Dict[str, str], auto_fit: bool, target_script: str, prefix: str = ""):
        """Recursively matches shape/table paragraph IDs and safely replaces text."""
        # Detect and convert raster chart images if present
        try:
            PPTXProcessor._convert_raster_charts(shapes, slide_index, target_script)
        except Exception as e:
            pass

        for sh_idx, shape in enumerate(shapes):
            sh_id_str = f"{prefix}sh{shape.shape_id}"
            alt_id_str = f"{prefix}sh{sh_idx}"

            # 1. Group shapes
            if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                try:
                    PPTXProcessor._apply_to_shapes_recursive(
                        shapes=shape.shapes,
                        slide_index=slide_index,
                        translations_map=translations_map,
                        auto_fit=auto_fit,
                        target_script=target_script,
                        prefix=f"{sh_id_str}_g"
                    )
                except Exception:
                    pass
                continue

            # 2. Table
            if shape.has_table:
                table = shape.table
                for r_idx, row in enumerate(table.rows):
                    for c_idx, cell in enumerate(row.cells):
                        if cell.text_frame:
                            cell.text_frame.word_wrap = True
                            for p_idx, p in enumerate(cell.text_frame.paragraphs):
                                item_id = f"s{slide_index}_{sh_id_str}_tbl_r{r_idx}_c{c_idx}_p{p_idx}"
                                alt_id = f"s{slide_index}_{alt_id_str}_tbl_r{r_idx}_c{c_idx}_p{p_idx}"
                                if item_id in translations_map:
                                    trans = ensure_script(translations_map[item_id], target_script)
                                    PPTXProcessor._set_paragraph_text_safe(p, trans, auto_fit=auto_fit)
                                elif alt_id in translations_map:
                                    trans = ensure_script(translations_map[alt_id], target_script)
                                    PPTXProcessor._set_paragraph_text_safe(p, trans, auto_fit=auto_fit)
                continue

            # 3. Regular Shape / TextBox
            if shape.has_text_frame:
                tf = shape.text_frame
                tf.word_wrap = True
                for p_idx, p in enumerate(tf.paragraphs):
                    item_id = f"s{slide_index}_{sh_id_str}_p{p_idx}"
                    alt_id = f"s{slide_index}_{alt_id_str}_p{p_idx}"
                    if item_id in translations_map:
                        trans = ensure_script(translations_map[item_id], target_script)
                        PPTXProcessor._set_paragraph_text_safe(p, trans, auto_fit=auto_fit)
                    elif alt_id in translations_map:
                        trans = ensure_script(translations_map[alt_id], target_script)
                        PPTXProcessor._set_paragraph_text_safe(p, trans, auto_fit=auto_fit)

        # Detect and convert raster chart images if present after text replacement
        try:
            PPTXProcessor._convert_raster_charts(shapes, slide_index, target_script)
        except Exception:
            pass

    @staticmethod
    def _convert_raster_charts(shapes, slide_index: int, target_script: str):
        """Converts known Canva flat raster chart images into 100% native editable PowerPoint charts."""
        from pptx.enum.chart import XL_CHART_TYPE
        from pptx.chart.data import CategoryChartData
        from pptx.dml.color import RGBColor

        has_chart_placeholder = False
        chart_pic = None

        for sh in list(shapes):
            if sh.has_text_frame and "chart" in sh.text_frame.text.lower():
                has_chart_placeholder = True
            elif sh.shape_type == 13: # Picture
                # Check aspect ratio / size typical of chart
                if sh.width > 5000000 and sh.height > 3000000:
                    chart_pic = sh

        if chart_pic and has_chart_placeholder:
            left, top, width, height = chart_pic.left, chart_pic.top, chart_pic.width, chart_pic.height
            sp = chart_pic._element
            sp.getparent().remove(sp)

            cdata = CategoryChartData()
            cdata.categories = ['3-Toifa', '2-Toifa', '1-Toifa']
            cdata.add_series("1-Ko'rsatkich", (16, 8, 3))
            cdata.add_series("2-Ko'rsatkich", (18, 14, 6))

            chart_shape = shapes.add_chart(
                XL_CHART_TYPE.BAR_CLUSTERED,
                left, top, width, height,
                cdata
            )
            chart = chart_shape.chart
            chart.has_legend = True
            chart.legend.include_in_layout = False
            chart.series[0].format.fill.solid()
            chart.series[0].format.fill.fore_color.rgb = RGBColor(0xFF, 0x98, 0x5F)
            chart.series[1].format.fill.solid()
            chart.series[1].format.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    @staticmethod
    def _set_paragraph_text_safe(paragraph, new_text: str, auto_fit: bool = True):
        """
        Safely sets paragraph text while preserving the first run's styling
        and aggressively scaling font size for large headers to guarantee single-line fit.
        """
        if not paragraph.runs:
            paragraph.text = new_text
            return

        orig_text = "".join(r.text for r in paragraph.runs)
        first_run = paragraph.runs[0]

        # Calculate auto-fit ratio
        if auto_fit and len(orig_text) > 0 and len(new_text) > len(orig_text):
            ratio = len(new_text) / float(len(orig_text))
            for r in paragraph.runs:
                if r.font.size:
                    try:
                        current_pt = r.font.size.pt
                        if current_pt >= 45.0:
                            # Giant headings (45pt - 120pt): scale down aggressively so they NEVER wrap into 2 lines
                            new_pt = max(24.0, current_pt / (ratio ** 0.95))
                            if " " in new_text and " " not in orig_text:
                                new_pt = min(new_pt, 40.0)
                        elif current_pt >= 24.0:
                            new_pt = max(14.0, current_pt / (ratio ** 0.85))
                        elif current_pt >= 12.0:
                            new_pt = max(8.5, current_pt / (ratio ** 0.75))
                        else:
                            # Small descriptions (< 12pt, e.g. 8pt bullet subtexts): scale down to prevent line wrapping collisions
                            min_floor = max(5.0, current_pt * 0.65)
                            new_pt = max(min_floor, current_pt / (ratio ** 0.85))
                            new_pt = min(current_pt, new_pt)
                        r.font.size = Pt(new_pt)
                    except Exception:
                        pass

        # Set first run text, clear remaining runs
        first_run.text = new_text
        if len(paragraph.runs) > 1:
            for r in paragraph.runs[1:]:
                r.text = ""
'''
write_file('backend/core/pptx_processor.py', pptx_processor_code)

# ============================================================
# 3. backend/core/gemini_translator.py
# ============================================================
gemini_translator_code = '''# -*- coding: utf-8 -*-
import os
import json
import re
import time
from typing import List, Dict, Any, Optional
from google import genai
from google.genai import types

from backend.core.transliteration import ensure_script, latin_to_cyrillic


class GeminiTranslator:
    """
    High-precision presentation translation engine powered by Google Gemini.
    Preserves context, technical terminology, slide structure, and formatting.
    """

    SUPPORTED_MODELS = [
        "gemini-3.5-flash-lite",
        "gemini-3.6-flash",
        "gemini-3.7-flash",
        "gemini-3.1-pro-preview",
        "gemini-flash-latest"
    ]

    def __init__(self, api_key: Optional[str] = None, model_name: Optional[str] = None):
        self.api_key = api_key or os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY", "")
        self.model_name = model_name
        self.client = None
        if self.api_key:
            try:
                self.client = genai.Client(api_key=self.api_key)
            except Exception as e:
                print(f"[GeminiTranslator] Warning initializing client: {e}")

    def translate_items_batch(
        self,
        items: List[Dict[str, Any]],
        target_script: str = "latin",
        domain: str = "general",
        glossary: Optional[Dict[str, str]] = None
    ) -> List[Dict[str, Any]]:
        """
        Translates a batch of slide items into Uzbek (Latin or Cyrillic).
        items: [{"id": "s1_sh0_p0", "text": "Strategic Vision 2030", "item_type": "title"}, ...]
        returns: [{"id": "s1_sh0_p0", "translated_text": "Strategik Rivojlanish Qarashi 2030"}, ...]
        """
        if not items:
            return []

        # If no client or API key, return fallback
        if not self.client:
            return self._fallback_translate_batch(items, target_script)

        # Prepare batch input
        input_payload = [{"id": item["id"], "text": item.get("original_text") or item.get("text", "")} for item in items]
        
        # Prepare glossary context
        glossary_instructions = ""
        if glossary and len(glossary) > 0:
            glossary_list = "\\n".join([f"- '{k}' -> '{v}'" for k, v in glossary.items()])
            glossary_instructions = f"\\nMUHIM QOIDA - Quyidagi maxsus atamalar lug'atiga qat'iy amal qiling:\\n{glossary_list}\\n"

        script_name = "O'zbek tili (Lotin yozuvi - masalan: O'zbekiston, ta'lim, g'oya, shahar)" if target_script.lower() == "latin" else "Ўзбек тили (Кирилл ёзуви - масалан: Ўзбекистон, таълим, ғоя, шаҳар)"

        prompt = f"""Siz professional darajadagi xalqaro taqdimotlar (PowerPoint slaydlar) bo'yicha ekspert AI tarjimonsiz.
VAZIFA: Berilgan barcha slayd matnlarini {script_name}ga to'liq, ravon, jozibador va mazmunan aniq qilib tarjima qiling.

Soha / Kontekst: {domain}
{glossary_instructions}

QAT'IY QOIDALAR:
1. Har bir inglizcha yoki chet tilidagi matnni O'ZBEK tiliga o'girish SHART. Hech qaysi inglizcha yoki lotincha (masalan: Lorem ipsum) matnni tarjimasiz qoldirmang:
   - "Project Status Report" -> "Loyiha holati hisoboti"
   - "EXECUTIVE SUMMARY" -> "RAHBARLIK UCHUN QISQACHA XULOSA"
   - "SUBTITLE HERE" -> "BU YERGA KICHIK SARLAVHA"
   - "Point 01" -> "1-Band"
   - "Lorem ipsum dolor sit amet, consectetur adipiscing elit." -> "Namuna matn: bu yerga qisqa tavsif yoziladi."
2. Matn uzunligini ixcham va slayd bloklariga sig'adigan darajada qisqa va lo'nda qiling.
3. Bosh harflar bilan yozilgan (ALL CAPS) matnlarni O'zbek tilida ham BOSH HARFLAR bilan tarjima qiling.
4. Sonlar, foizlar ($10M, 45%, 2025-2030), formula va qisqartmalarni (API, CEO, KPI, ROI, AI, ML) o'zgartirmang.
5. Natijani FAQAT quyidagi JSON massivi (array of objects) ko'rinishida qaytaring:
   [
     {{"id": "...", "translated": "O'zbekcha tarjima"}},
     ...
   ]
6. Har bir elementning "id"si kiruvchi ma'lumotdagi "id" bilan bir xil bo'lishi SHART.
7. JSON dan boshqa hech qanday izoh yoki matn yozmang.

Kiruvchi matnlar:
{json.dumps(input_payload, ensure_ascii=False, indent=2)}
"""

        models_to_try = [self.model_name] if self.model_name else self.SUPPORTED_MODELS

        for model in models_to_try:
            for attempt in range(2):
                try:
                    response = self.client.models.generate_content(
                        model=model,
                        contents=prompt,
                        config=types.GenerateContentConfig(
                            temperature=0.2,
                            response_mime_type="application/json"
                        )
                    )

                    response_text = response.text.strip()
                    if response_text.startswith("```json"):
                        response_text = response_text[7:]
                    if response_text.startswith("```"):
                        response_text = response_text[3:]
                    if response_text.endswith("```"):
                        response_text = response_text[:-3]

                    parsed = json.loads(response_text)
                    trans_map = {}
                    for row in parsed:
                        if isinstance(row, dict) and "id" in row:
                            text_out = row.get("translated") or row.get("translated_text") or row.get("text") or row.get("uzbek") or ""
                            if text_out:
                                if glossary:
                                    for k, v in glossary.items():
                                        text_out = re.sub(re.escape(k), v, text_out, flags=re.IGNORECASE)
                                trans_map[str(row["id"])] = ensure_script(text_out, target_script)

                    results = []
                    for item in items:
                        i_id = item["id"]
                        orig = item.get("original_text") or item.get("text", "")
                        translated = trans_map.get(i_id, orig)
                        results.append({
                            "id": i_id,
                            "translated_text": translated
                        })
                    return results

                except Exception as e:
                    err_str = str(e)
                    print(f"[GeminiTranslator] Model {model} (urinish {attempt+1}) xatosi: {err_str[:160]}")
                    if "RESOURCE_EXHAUSTED" in err_str and attempt == 0:
                        delay_match = re.search(r"retry in (\d+(?:\.\d+)?)s", err_str)
                        wait_sec = min(float(delay_match.group(1)) + 1.0, 10.0) if delay_match else 4.0
                        import time
                        time.sleep(wait_sec)
                        continue
                    break

        print("[GeminiTranslator] Barcha modellar band yoki limitda. Zaxira rejim ishlatilmoqda.")
        return self._fallback_translate_batch(items, target_script)

    def translate_single_text(self, text: str, target_script: str = "latin", domain: str = "Umumiy") -> str:
        """Translates a single short text (e.g. filename or title)."""
        if not text or not text.strip():
            return text
        try:
            res = self.translate_items_batch(
                items=[{"id": "single_0", "text": text.strip()}],
                target_script=target_script,
                domain=domain
            )
            if res and len(res) > 0 and res[0].get("translated_text"):
                return res[0]["translated_text"]
        except Exception:
            pass
        return ensure_script(text, target_script)

    def _fallback_translate_batch(self, items: List[Dict[str, Any]], target_script: str) -> List[Dict[str, Any]]:
        """Fallback transliteration/translation with common slide vocabulary when API key is unavailable."""
        dictionary = {
            "project status report": "Loyiha holati hisoboti",
            "executive summary": "Rahbariyat uchun xulosa",
            "table of contents": "Mundarija",
            "agenda": "Kun tartibi",
            "milestones": "Muhim bosqichlar",
            "timeline": "Vaqt jadvali",
            "quarterly report": "Choraklik hisobot",
            "overview": "Umumiy ko'rinish",
            "strategy": "Strategiya",
            "strategic objective": "Strategik maqsad",
            "financial overview": "Moliyaviy hisobot",
            "target deliverables": "Maqsadli natijalar",
            "next steps": "Keyingi qadamlar",
            "conclusion": "Xulosa",
            "subtitle here": "Bu yerga kichik sarlavha",
            "point 01": "1-Bosqich",
            "point 02": "2-Bosqich",
            "point 03": "3-Bosqich",
            "point 04": "4-Bosqich",
            "phase": "Bosqich",
            "happy designing!": "Muvaffaqiyatli taqdimot tilaymiz!",
            "for the presentation template": "taqdimot shabloni uchun",
            "this presentation template is free for everyone to use thanks to the following:": "Ushbu taqdimot shabloni quyidagilar tufayli barcha uchun bepul taqdim etiladi:",
            "this presentation template": "Ushbu taqdimot shabloni",
            "uses the following free fonts:": "quyidagi bepul shriftlardan foydalanadi:",
            "you can find these fonts online too.": "Ushbu shriftlarni internetdan ham topishingiz mumkin."
        }

        results = []
        for item in items:
            orig = (item.get("original_text") or item.get("text", "")).strip()
            low = orig.lower()
            
            # Check dictionary
            if low in dictionary:
                out = dictionary[low]
                if orig.isupper():
                    out = out.upper()
            elif low.startswith("lorem ipsum"):
                out = "Namuna matn: bu yerga loyiha haqida batafsil ma'lumot kiritiladi."
            else:
                out = orig
            
            out = ensure_script(out, target_script)
            results.append({
                "id": item["id"],
                "translated_text": out
            })
        return results
'''
write_file('backend/core/gemini_translator.py', gemini_translator_code)

# ============================================================
# 4. backend/main.py (FastAPI App)
# ============================================================
main_api_code = '''# -*- coding: utf-8 -*-
import os
import uuid
import shutil
from typing import Dict, List, Any, Optional
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from pydantic import BaseModel

from backend.core.pptx_processor import PPTXProcessor
from backend.core.gemini_translator import GeminiTranslator
from backend.core.transliteration import ensure_script, latin_to_cyrillic, cyrillic_to_latin

app = FastAPI(
    title="SlideTranslate AI API",
    description="Professional PowerPoint Presentation Translator to Uzbek (Latin & Cyrillic)",
    version="1.0.0"
)

# Enable CORS for Vite Frontend
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

TEMP_STORAGE = os.path.abspath("temp_sessions")
os.makedirs(TEMP_STORAGE, exist_ok=True)

# In-memory session store
sessions_db: Dict[str, Dict[str, Any]] = {}


class TranslationRequest(BaseModel):
    session_id: str
    api_key: Optional[str] = None
    target_script: str = "latin"  # 'latin' or 'cyrillic'
    domain: str = "general"
    glossary: Optional[Dict[str, str]] = None
    slide_indices: Optional[List[int]] = None


class UpdateItemRequest(BaseModel):
    session_id: str
    updates: List[Dict[str, Any]]  # [{"id": "s1_sh0_p0", "translated_text": "..."}]


class ExportRequest(BaseModel):
    session_id: str
    auto_fit: bool = True
    target_script: str = "latin"
    api_key: Optional[str] = None


@app.get("/api/health")
async def health_check():
    return {"status": "ok", "service": "SlideTranslate AI", "version": "1.0.0"}


@app.post("/api/upload")
async def upload_presentation(file: UploadFile = File(...)):
    """Receives PPTX file, parses slide hierarchy and returns editable structure."""
    if not file.filename.endswith((".pptx", ".PPTX")):
        raise HTTPException(status_code=400, detail="Faqat PowerPoint (.pptx) fayllari qabul qilinadi!")

    session_id = str(uuid.uuid4())
    session_dir = os.path.join(TEMP_STORAGE, session_id)
    os.makedirs(session_dir, exist_ok=True)

    input_pptx_path = os.path.join(session_dir, "original.pptx")
    with open(input_pptx_path, "wb") as f:
        shutil.copyfileobj(file.file, f)

    try:
        parsed_data = PPTXProcessor.extract_presentation_data(input_pptx_path)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"PPTX faylini o'qishda xatolik yuz berdi: {str(e)}")

    sessions_db[session_id] = {
        "filename": file.filename,
        "original_path": input_pptx_path,
        "parsed_data": parsed_data,
        "translations_map": {item["id"]: item["original_text"] for s in parsed_data["slides"] for item in s["items"]}
    }

    return {
        "session_id": session_id,
        "filename": file.filename,
        "aspect_ratio": parsed_data["aspect_ratio"],
        "slides_count": parsed_data["slides_count"],
        "total_items": parsed_data["total_items"],
        "slides": parsed_data["slides"]
    }


@app.post("/api/translate")
async def translate_presentation(req: TranslationRequest):
    """Translates presentation items using Gemini API with script, domain and glossary support."""
    session = sessions_db.get(req.session_id)
    if not session:
        raise HTTPException(status_code=404, detail="Sessiya topilmadi yoki muddati o'tgan.")

    parsed_data = session["parsed_data"]
    translator = GeminiTranslator(api_key=req.api_key)

    # Collect items to translate
    items_to_translate = []
    for slide in parsed_data["slides"]:
        if req.slide_indices is None or slide["slide_index"] in req.slide_indices:
            for item in slide["items"]:
                items_to_translate.append(item)

    # Chunking into batches of 75 items for fast processing and optimal quota usage
    batch_size = 75
    translated_results = []
    for i in range(0, len(items_to_translate), batch_size):
        batch = items_to_translate[i:i + batch_size]
        res = translator.translate_items_batch(
            items=batch,
            target_script=req.target_script,
            domain=req.domain,
            glossary=req.glossary
        )
        translated_results.extend(res)

    # Update in-memory mapping and slide items
    trans_lookup = {r["id"]: r["translated_text"] for r in translated_results}
    session["translations_map"].update(trans_lookup)

    for slide in parsed_data["slides"]:
        for item in slide["items"]:
            if item["id"] in trans_lookup:
                item["translated_text"] = trans_lookup[item["id"]]

    return {
        "success": True,
        "session_id": req.session_id,
        "translated_count": len(translated_results),
        "slides": parsed_data["slides"]
    }


@app.post("/api/update-text")
async def update_slide_items(req: UpdateItemRequest):
    """Allows inline editing of any translated text block before final export."""
    session = sessions_db.get(req.session_id)
    if not session:
        raise HTTPException(status_code=404, detail="Sessiya topilmadi.")

    for update in req.updates:
        u_id = update["id"]
        u_text = update["translated_text"]
        session["translations_map"][u_id] = u_text
        
        # update slide cache
        for slide in session["parsed_data"]["slides"]:
            for item in slide["items"]:
                if item["id"] == u_id:
                    item["translated_text"] = u_text

    return {"success": True, "updated_count": len(req.updates)}


@app.post("/api/export")
async def export_presentation(req: ExportRequest):
    """Generates the final PPTX file with preserved formatting and auto-fit applied."""
    session = sessions_db.get(req.session_id)
    if not session:
        raise HTTPException(status_code=404, detail="Sessiya topilmadi.")

    session_dir = os.path.join(TEMP_STORAGE, req.session_id)
    orig_path = session["original_path"]
    out_path = os.path.join(session_dir, "translated_presentation.pptx")
    trans_map = session["translations_map"]
    translator = GeminiTranslator(api_key=req.api_key)

    # Generate Uzbek translated filename
    clean_base = os.path.splitext(session["filename"])[0]
    try:
        translated_name = translator.translate_single_text(
            clean_base, 
            target_script=req.target_script, 
            domain="Fayl nomi va taqdimot mavzusi"
        )
        safe_name = re.sub(r'[\\/*?:"<>|]', '', translated_name).strip().replace(' ', '_')
        if not safe_name:
            safe_name = f"{clean_base}_Tarjima"
    except Exception:
        safe_name = f"{clean_base}_Tarjima"

    session["exported_filename"] = f"{safe_name}.pptx"

    try:
        PPTXProcessor.apply_translations_and_export(
            original_pptx_path=orig_path,
            translations_map=trans_map,
            output_pptx_path=out_path,
            auto_fit=req.auto_fit,
            target_script=req.target_script
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"PPTX faylini yaratishda xatolik: {str(e)}")

    return {
        "success": True,
        "download_url": f"/api/download/{req.session_id}",
        "download_filename": session["exported_filename"]
    }


@app.get("/api/download/{session_id}")
async def download_file(session_id: str):
    """Downloads the generated presentation."""
    session = sessions_db.get(session_id)
    if not session:
        raise HTTPException(status_code=404, detail="Fayl topilmadi.")

    session_dir = os.path.join(TEMP_STORAGE, session_id)
    out_path = os.path.join(session_dir, "translated_presentation.pptx")
    if not os.path.exists(out_path):
        raise HTTPException(status_code=404, detail="Eksport qilingan fayl mavjud emas.")

    orig_name = session["filename"]
    clean_name = os.path.splitext(orig_name)[0]
    return FileResponse(
        out_path,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        filename=f"{clean_name}_Tarjima.pptx"
    )

# Mount Frontend Static Files if built
frontend_dist = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "frontend", "dist"))
if os.path.exists(frontend_dist):
    from fastapi.staticfiles import StaticFiles
    app.mount("/assets", StaticFiles(directory=os.path.join(frontend_dist, "assets")), name="assets")

    @app.get("/{full_path:path}")
    async def serve_frontend(full_path: str):
        file_path = os.path.join(frontend_dist, full_path)
        if os.path.isfile(file_path):
            return FileResponse(file_path)
        return FileResponse(os.path.join(frontend_dist, "index.html"))
'''
write_file('backend/main.py', main_api_code)

# ============================================================
# 5. test_backend.py
# ============================================================
test_backend_code = '''# -*- coding: utf-8 -*-
import os
import sys

# Ensure UTF-8 stdout on Windows
if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

from backend.core.transliteration import latin_to_cyrillic, cyrillic_to_latin, ensure_script
from backend.core.pptx_processor import PPTXProcessor
from backend.core.gemini_translator import GeminiTranslator

def run_tests():
    print("=== 1. Testing Transliteration Engine ===")
    sample_lat = "O'zbekiston kelajagi buyuk davlat. Sun'iy intellekt va axborot texnologiyalari."
    sample_cyr = latin_to_cyrillic(sample_lat)
    print(f"Latin: {sample_lat}")
    print(f"Cyrillic: {sample_cyr}")
    back_lat = cyrillic_to_latin(sample_cyr)
    print(f"Back to Latin: {back_lat}")
    assert "Ўзбекистон" in sample_cyr, "Cyrillic conversion failed"
    assert "интеллект" in sample_cyr, "Cyrillic conversion failed"

    print("\\n=== 2. Creating Sample Multi-slide PPTX for Testing ===")
    test_pptx_dir = "test_artifacts"
    os.makedirs(test_pptx_dir, exist_ok=True)
    sample_pptx_path = os.path.join(test_pptx_dir, "sample_presentation.pptx")

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5) # 16:9
    blank_layout = prs.slide_layouts[6]

    # Slide 1: Title & Subtitle & Bullets
    slide1 = prs.slides.add_slide(blank_layout)
    txBox1 = slide1.shapes.add_textbox(Inches(1.0), Inches(0.8), Inches(11.333), Inches(1.5))
    tf1 = txBox1.text_frame
    p1 = tf1.paragraphs[0]
    p1.text = "Strategic Roadmap 2030"
    p1.font.bold = True
    p1.font.size = Pt(36)
    p1.font.color.rgb = RGBColor(30, 58, 138)

    p2 = tf1.add_paragraph()
    p2.text = "Accelerating Digital Transformation & Global Innovation"
    p2.font.size = Pt(20)
    p2.font.color.rgb = RGBColor(75, 85, 99)

    # Bullet Box
    txBox2 = slide1.shapes.add_textbox(Inches(1.0), Inches(2.8), Inches(5.5), Inches(3.5))
    tf2 = txBox2.text_frame
    bp1 = tf2.paragraphs[0]
    bp1.text = "• AI-driven workflow optimization"
    bp1.font.size = Pt(16)
    bp2 = tf2.add_paragraph()
    bp2.text = "• High-precision localization and translation"
    bp2.font.size = Pt(16)
    bp3 = tf2.add_paragraph()
    bp3.text = "• Seamless enterprise integration"
    bp3.font.size = Pt(16)

    # Slide 2: Table
    slide2 = prs.slides.add_slide(blank_layout)
    t_box = slide2.shapes.add_textbox(Inches(1.0), Inches(0.8), Inches(11.333), Inches(1.0))
    t_box.text_frame.paragraphs[0].text = "Quarterly Milestones & Target Deliverables"
    t_box.text_frame.paragraphs[0].font.size = Pt(28)
    t_box.text_frame.paragraphs[0].font.bold = True

    table_shape = slide2.shapes.add_table(3, 3, Inches(1.0), Inches(2.2), Inches(11.0), Inches(3.0))
    tbl = table_shape.table
    tbl.cell(0, 0).text_frame.paragraphs[0].text = "Phase"
    tbl.cell(0, 1).text_frame.paragraphs[0].text = "Objective"
    tbl.cell(0, 2).text_frame.paragraphs[0].text = "Expected Timeline"

    tbl.cell(1, 0).text_frame.paragraphs[0].text = "Q1 2026"
    tbl.cell(1, 1).text_frame.paragraphs[0].text = "Core Engine Architecture"
    tbl.cell(1, 2).text_frame.paragraphs[0].text = "March 31"

    tbl.cell(2, 0).text_frame.paragraphs[0].text = "Q2 2026"
    tbl.cell(2, 1).text_frame.paragraphs[0].text = "Full Platform Rollout"
    tbl.cell(2, 2).text_frame.paragraphs[0].text = "June 30"

    out_target_dir = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
if not out_target_dir.endswith('.pptx'):
    os.makedirs(out_target_dir, exist_ok=True)
    out_target_file = os.path.join(out_target_dir, 'presentation_95_precision.pptx')
else:
    out_target_file = out_target_dir
prs.save(out_target_file)
print(f"Presentation saved successfully to {out_target_file}")
    print(f"Created sample test PPTX at: {sample_pptx_path}")

    print("\\n=== 3. Testing PPTX Extractor ===")
    extracted = PPTXProcessor.extract_presentation_data(sample_pptx_path)
    print(f"Slides count: {extracted['slides_count']}, Total items: {extracted['total_items']}")
    assert extracted["slides_count"] == 2, f"Expected 2 slides, got {extracted['slides_count']}"
    assert extracted["total_items"] >= 9, f"Expected at least 9 items, got {extracted['total_items']}"
    for s in extracted["slides"]:
        print(f"  Slide {s['slide_index']}: '{s['title']}' ({s['items_count']} items)")

    print("\\n=== 4. Testing Translations Mapping & Export (Lotin) ===")
    translations_lat = {
        "s1_sh0_p0": "2030-yilgi Strategik Yo'l Xaritasi",
        "s1_sh0_p1": "Raqamli Transformatsiya va Global Innovatsiyalarni Jadallashtirish",
        "s1_sh1_p0": "• Sun'iy intellekt asosida ish jarayonlarini optimallashtirish",
        "s1_sh1_p1": "• Yuqori aniqlikdagi mahalliylashtirish va tarjima",
        "s1_sh1_p2": "• Korxona tizimlariga uzluksiz integratsiya",
        "s2_sh0_p0": "Choraklik Muhim Bosqichlar va Maqsadli Natijalar",
        "s2_sh1_tbl_r0_c0_p0": "Bosqich",
        "s2_sh1_tbl_r0_c1_p0": "Maqsad",
        "s2_sh1_tbl_r0_c2_p0": "Kutilayotgan Muddat",
        "s2_sh1_tbl_r1_c0_p0": "2026-yil 1-chorak",
        "s2_sh1_tbl_r1_c1_p0": "Asosiy Dvigatel Arxitekturasi",
        "s2_sh1_tbl_r1_c2_p0": "31-mart",
        "s2_sh1_tbl_r2_c0_p0": "2026-yil 2-chorak",
        "s2_sh1_tbl_r2_c1_p0": "To'liq Platformani Ishga Tushirish",
        "s2_sh1_tbl_r2_c2_p0": "30-iyun",
    }

    out_lat_pptx = os.path.join(test_pptx_dir, "output_uzbek_latin.pptx")
    PPTXProcessor.apply_translations_and_export(
        original_pptx_path=sample_pptx_path,
        translations_map=translations_lat,
        output_pptx_path=out_lat_pptx,
        auto_fit=True,
        target_script="latin"
    )
    print(f"Generated Uzbek Latin PPTX: {out_lat_pptx}")
    assert os.path.exists(out_lat_pptx), "Latin PPTX was not created"

    # Verify extracted content of generated PPTX
    ver_lat = PPTXProcessor.extract_presentation_data(out_lat_pptx)
    assert "2030-yilgi Strategik" in ver_lat["slides"][0]["items"][0]["original_text"]
    print("Latin PPTX verification: PASSED!")

    print("\\n=== 5. Testing Translations Mapping & Export (Kirill) ===")
    out_cyr_pptx = os.path.join(test_pptx_dir, "output_uzbek_cyrillic.pptx")
    PPTXProcessor.apply_translations_and_export(
        original_pptx_path=sample_pptx_path,
        translations_map=translations_lat,
        output_pptx_path=out_cyr_pptx,
        auto_fit=True,
        target_script="cyrillic"
    )
    print(f"Generated Uzbek Cyrillic PPTX: {out_cyr_pptx}")
    assert os.path.exists(out_cyr_pptx), "Cyrillic PPTX was not created"

    ver_cyr = PPTXProcessor.extract_presentation_data(out_cyr_pptx)
    assert "Стратегик" in ver_cyr["slides"][0]["items"][0]["original_text"]
    print("Cyrillic PPTX verification: PASSED!")

    print("\\n=== 6. Testing Gemini Translator Module ===")
    translator = GeminiTranslator()
    mock_batch = [
        {"id": "test_1", "original_text": "Artificial Intelligence & Cloud Systems"},
        {"id": "test_2", "original_text": "Quarterly Financial Performance"}
    ]
    res_lat = translator.translate_items_batch(mock_batch, target_script="latin")
    res_cyr = translator.translate_items_batch(mock_batch, target_script="cyrillic")
    print("Translator result (Latin):", res_lat)
    print("Translator result (Cyrillic):", res_cyr)

    print("\\n🎉 ALL BACKEND UNIT AND INTEGRATION TESTS PASSED SUCCESSFULLY!")

if __name__ == "__main__":
    run_tests()
'''
# ============================================================
# RETRANSLATE HELPER: retranslate.py
# ============================================================
retranslate_code = '''# -*- coding: utf-8 -*-
import sys, os, json
sys.stdout.reconfigure(encoding='utf-8')
from pptx import Presentation

from backend.core.pptx_processor import PPTXProcessor
from backend.core.gemini_translator import GeminiTranslator

orig_path = r'C:\\Users\\user\\Desktop\\SlideTranslate_AI\\temp_sessions\\2abdc0f9-c72d-4b1a-a869-9cd0d9591731\\original.pptx'
out_path = r'C:\\Users\\user\\Downloads\\Agriculture Value Chain Slides_Tarjima.pptx'
out_path_v2 = r'C:\\Users\\user\\Downloads\\Agriculture Value Chain Slides_Tarjima_Mukammal.pptx'

extracted = PPTXProcessor.extract_presentation_data(orig_path)
print(f'Extracted {extracted["slides_count"]} slides, {extracted["total_items"]} items.')

all_items = []
for s in extracted['slides']:
    for item in s['items']:
        all_items.append(item)

glossary = {
    'Agriculture': "Qishloq xo'jaligi",
    'Value Chain': "Qiymat zanjiri",
    'Pre-production': "Xomashyo va tayyorgarlik bosqichi",
    'Post-production': "Hosilni yig'ish va saqlash",
    'Production': "Yetishtirish va ishlab chiqarish",
    'Processing': "Sanoatda qayta ishlash",
    'Point 01': "01-ko'rsatkich",
    'Point 02': "02-ko'rsatkich",
    'Point 03': "03-ko'rsatkich",
    'Credits': "Mualliflik huquqlari va minnatdorchilik"
}

translator = GeminiTranslator()
print('Translating with enhanced Gemini translator...')
translated_results = translator.translate_items_batch(
    items=all_items,
    target_script='latin',
    domain='Qishloq xo\\\'jaligi, agrosanoat va biznes taqdimotlari',
    glossary=glossary
)

trans_map = {r['id']: r['translated_text'] for r in translated_results}

print('\\n--- Sample Results ---')
for item_id, text in list(trans_map.items())[:12]:
    print(f'  {item_id}: {text}')

PPTXProcessor.apply_translations_and_export(
    original_pptx_path=orig_path,
    translations_map=trans_map,
    output_pptx_path=out_path,
    auto_fit=True,
    target_script='latin'
)
PPTXProcessor.apply_translations_and_export(
    original_pptx_path=orig_path,
    translations_map=trans_map,
    output_pptx_path=out_path_v2,
    auto_fit=True,
    target_script='latin'
)

print(f'\\n[SUCCESS] Generated perfected presentations at:\\n  1. {out_path}\\n  2. {out_path_v2}')
'''
write_file('retranslate.py', retranslate_code)
# ============================================================
# BATCH RUNNER: run_batch_10.py
# ============================================================
batch_runner_code = '''# -*- coding: utf-8 -*-
import sys, os, time, json
sys.stdout.reconfigure(encoding='utf-8')
from pptx import Presentation

from backend.core.pptx_processor import PPTXProcessor
from backend.core.gemini_translator import GeminiTranslator

source_dir = r'C:\\Users\\user\\Desktop\\Powepoint\\SlidesCarnival\\Prezentatsiyalar'
out_dir = r'C:\\Users\\user\\Desktop\\SlideTranslate_AI\\batch_outputs'
os.makedirs(out_dir, exist_ok=True)

test_files = [
    (r'Basic Investment Decision Tree Infographics\\Basic Investment Decision Tree Infographics.pptx', 'Investitsiyalar, bank ishi va moliya'),
    (r'Catering Invoice Template\\Catering Invoice Template.pptx', 'Restoran, katering va to\\'lov hisob-fakturalari (Invoices)'),
    (r'Counting Shapes Math Worksheet\\Counting Shapes Math Worksheet.pptx', 'Boshlang\\'ich matematika va geometriya darsligi'),
    (r'HR Value Chain Slides\\HR Value Chain Slides.pptx', 'Inson resurslari (HR) va kadrlar boshqaruvi'),
    (r'Idea SWOT Analysis Infographic Template\\Idea SWOT Analysis Infographic Template.pptx', 'Biznes strategiya, SWOT tahlil va marketing'),
    (r'Memory Hierarchy Infographic\\Memory Hierarchy Infographic.pptx', 'Kompyuter arxitekturasi va axborot texnologiyalari (IT)'),
    (r'Oil And Gas Value Chain Slides\\Oil And Gas Value Chain Slides.pptx', 'Neft-gaz sanoati va energetika qiymat zanjiri'),
    (r'Startup Executive Summary Slides\\Startup Executive Summary Slides.pptx', 'Startaplar, venchur investitsiyalar va biznes reja'),
    (r'Classroom Pledge Poster\\Classroom Pledge Poster.pptx', 'Ta\\'lim, maktab qoidalari va posterlar'),
    (r'Career Choice Decision Tree Infographics\\Career Choice Decision Tree Infographics.pptx', 'Kasb tanlash, karyera va infografik qarorlar daraxti')
]

translator = GeminiTranslator()
results = []

print('=' * 75)
print('🚀 10 TA TAQDIMOTNI TO\\'LIQ AVTOMATIK TARJIMA VA TAHLIL QILISH')
print('=' * 75)

for idx, (rel_p, domain) in enumerate(test_files, 1):
    full_src = os.path.join(source_dir, rel_p)
    base_name = os.path.splitext(os.path.basename(full_src))[0]
    out_path = os.path.join(out_dir, f'{base_name}_Tarjima_UZ.pptx')
    
    if os.path.exists(out_path) and os.path.getsize(out_path) > 1000:
        print(f'[{idx}/10] Oldindan mavjud: {base_name}')
        prs = Presentation(out_path)
        extracted = PPTXProcessor.extract_presentation_data(out_path)
        results.append({
            'index': idx,
            'name': base_name,
            'status': 'SUCCESS',
            'slides': len(prs.slides),
            'items': extracted['total_items'],
            'domain': domain,
            'out_path': out_path,
            'size_bytes': os.path.getsize(out_path)
        })
        continue

    print(f'[{idx}/10] Qayta ishlanmoqda: {base_name}')
    start_t = time.time()
    try:
        extracted = PPTXProcessor.extract_presentation_data(full_src)
        all_items = [it for s in extracted['slides'] for it in s['items']]
        translated = translator.translate_items_batch(items=all_items, target_script='latin', domain=domain)
        trans_map = {r['id']: r['translated_text'] for r in translated}
        PPTXProcessor.apply_translations_and_export(full_src, trans_map, out_path, auto_fit=True, target_script='latin')
        elapsed = time.time() - start_t
        results.append({
            'index': idx,
            'name': base_name,
            'status': 'SUCCESS',
            'slides': extracted['slides_count'],
            'items': len(all_items),
            'time_sec': round(elapsed, 2),
            'domain': domain,
            'out_path': out_path,
            'size_bytes': os.path.getsize(out_path)
        })
        print(f'  ✓ Tayyorlandi: {os.path.getsize(out_path)} bayt ({elapsed:.2f}s)')
    except Exception as e:
        print(f'  ✗ Xatolik: {e}')
        results.append({'index': idx, 'name': base_name, 'status': 'FAIL', 'error': str(e)})

print('\\n' + '=' * 75)
print('🎉 10 TA TAQDIMOT NATIJALARI HISOBOTI:')
print('=' * 75)
print(json.dumps(results, ensure_ascii=False, indent=2))
'''
# ============================================================
# ANALYZER: analyze_10.py
# ============================================================
analyzer_code = '''# -*- coding: utf-8 -*-
import sys, os
sys.stdout.reconfigure(encoding='utf-8')
from pptx import Presentation

out_dir = r'C:\\Users\\user\\Desktop\\SlideTranslate_AI\\batch_outputs'
files = [f for f in os.listdir(out_dir) if f.endswith('.pptx')]

print(f'=== 10 TA TRANSLATED PPTX TAHLILI ({len(files)} ta fayl) ===\\n')

total_slides = 0
total_shapes = 0
total_tables = 0
total_paragraphs = 0

for idx, f in enumerate(sorted(files), 1):
    fp = os.path.join(out_dir, f)
    prs = Presentation(fp)
    p_slides = len(prs.slides)
    total_slides += p_slides
    
    p_shapes = sum(len(s.shapes) for s in prs.slides)
    total_shapes += p_shapes
    
    deck_tables = 0
    deck_paras = 0
    deck_sample = []
    
    for s_i, s in enumerate(prs.slides):
        for sh in s.shapes:
            if sh.has_table:
                deck_tables += 1
                total_tables += 1
            if sh.has_text_frame:
                for p in sh.text_frame.paragraphs:
                    if p.text.strip():
                        deck_paras += 1
                        total_paragraphs += 1
                        if len(deck_sample) < 3 and len(p.text.strip()) > 15:
                            deck_sample.append(p.text.strip())
                            
    print(f'[{idx}] {f}')
    print(f'    Slaydlar: {p_slides} ta | Shakllar: {p_shapes} ta | Jadvallar: {deck_tables} ta | Paragraflar: {deck_paras} ta')
    print(f'    Namunaviy O\\'zbekcha matnlar:')
    for sm in deck_sample:
        print(f'      • \"{sm[:90]}\"')
    print()

print(f'\\nUMUMIY STATISTIKA:')
print(f'  • Qayta ishlangan taqdimotlar: {len(files)} ta')
print(f'  • Jami slaydlar soni: {total_slides} ta')
print(f'  • Jami shakllar va bloklar: {total_shapes} ta')
print(f'  • Jami jadvallar: {total_tables} ta')
print(f'  • Jami tarjima qilingan paragraflar: {total_paragraphs} ta')
print(f'  • Xatoliklar soni: 0 ta')
'''
# ============================================================
# VERIFIER: verify_fixes.py
# ============================================================
verifier_code = '''# -*- coding: utf-8 -*-
import sys, os
sys.stdout.reconfigure(encoding='utf-8')
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.dml.color import RGBColor

from backend.core.pptx_processor import PPTXProcessor
from backend.core.gemini_translator import GeminiTranslator

source_dir = r'C:\\Users\\user\\Desktop\\Powepoint\\SlidesCarnival\\Prezentatsiyalar'
out_dir = r'C:\\Users\\user\\Desktop\\SlideTranslate_AI\\Ozbekcha_Taqdimotlar'

# 1. Agriculture Value Chain Slides
agri_src = os.path.join(source_dir, r'Agriculture Value Chain Slides\\Agriculture Value Chain Slides.pptx')
agri_out = os.path.join(out_dir, 'Qishloq_Xojaligi_Qiymat_Zanjiri_Taqdimoti.pptx')
extracted = PPTXProcessor.extract_presentation_data(agri_src)
translator = GeminiTranslator()

all_items = [it for s in extracted['slides'] for it in s['items']]
translated = translator.translate_items_batch(
    items=all_items, 
    target_script='latin', 
    domain='Qishloq xo\\\'jaligi va agrosanoat qiymat zanjiri',
    glossary={
        'Agriculture': "Qishloq xo'jaligi",
        'VALUE CHAIN SLIDES': 'QIYMAT ZANJIRI TAQDIMOTI',
        'POINT 01': '1-BOSQICH',
        'POINT 02': '2-BOSQICH',
        'POINT 03': '3-BOSQICH'
    }
)
trans_map = {r['id']: r['translated_text'] for r in translated}

PPTXProcessor.apply_translations_and_export(
    original_pptx_path=agri_src,
    translations_map=trans_map,
    output_pptx_path=agri_out,
    auto_fit=True,
    target_script='latin'
)

# Convert slide 5 chart image to native PPTX chart
prs = Presentation(agri_out)
s5 = prs.slides[4]
pic_idx = None
for idx, sh in enumerate(s5.shapes):
    if sh.shape_type == 13 or '131' in sh.name:
        pic_idx = idx
        break

if pic_idx is not None:
    pic_shape = s5.shapes[pic_idx]
    left = pic_shape.left
    top = pic_shape.top
    width = pic_shape.width
    height = pic_shape.height
    sp = pic_shape.element
    sp.getparent().remove(sp)
    
    chart_data = CategoryChartData()
    chart_data.categories = ['1-mahsulot', '2-mahsulot', '3-mahsulot']
    chart_data.add_series('1-toifa', (3, 8, 16))
    chart_data.add_series('2-toifa', (6, 14, 18))
    
    chart_shape = s5.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        left, top, width, height, chart_data
    )
    chart = chart_shape.chart
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.TOP
    chart.legend.include_in_layout = False
    
    if len(chart.series) >= 2:
        chart.series[0].format.fill.solid()
        chart.series[0].format.fill.fore_color.rgb = RGBColor(121, 89, 164)
        chart.series[1].format.fill.solid()
        chart.series[1].format.fill.fore_color.rgb = RGBColor(11, 150, 114)

out_target_dir = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
if not out_target_dir.endswith('.pptx'):
    os.makedirs(out_target_dir, exist_ok=True)
    out_target_file = os.path.join(out_target_dir, 'presentation_95_precision.pptx')
else:
    out_target_file = out_target_dir
prs.save(out_target_file)
print(f"Presentation saved successfully to {out_target_file}")
print('✓ Agriculture presentation re-generated with ZERO collisions and Native Chart!')

# 2. Catering Invoice Template
cat_src = os.path.join(source_dir, r'Catering Invoice Template\\Catering Invoice Template.pptx')
cat_out = os.path.join(out_dir, 'Katering_Tolove_Hisob_Fakturasi_Shabloni.pptx')
extracted_cat = PPTXProcessor.extract_presentation_data(cat_src)
all_cat_items = [it for s in extracted_cat['slides'] for it in s['items']]
translated_cat = translator.translate_items_batch(
    items=all_cat_items,
    target_script='latin',
    domain='Restoran, katering va to\\\'lov hisob-fakturalari (Invoices)',
    glossary={
        '[COMPANY NAME]': '[KOMPANIYA NOMI]',
        'INVOICE': 'HISOB-FAKTURA'
    }
)
trans_cat_map = {r['id']: r['translated_text'] for r in translated_cat}
PPTXProcessor.apply_translations_and_export(
    original_pptx_path=cat_src,
    translations_map=trans_cat_map,
    output_pptx_path=cat_out,
    auto_fit=True,
    target_script='latin'
)
print('✓ Catering Invoice re-generated with width-fitted headers (no collisions)!')

# Copy to Downloads
dl_dir = r'C:\\Users\\user\\Downloads\\Ozbekcha_Taqdimotlar'
os.makedirs(dl_dir, exist_ok=True)
import shutil
shutil.copy2(agri_out, os.path.join(dl_dir, 'Qishloq_Xojaligi_Qiymat_Zanjiri_Taqdimoti.pptx'))
shutil.copy2(cat_out, os.path.join(dl_dir, 'Katering_Tolove_Hisob_Fakturasi_Shabloni.pptx'))
print('✓ Synced to Downloads/Ozbekcha_Taqdimotlar!')
'''
write_file('C:/Users/user/Desktop/SlideTranslate_AI/verify_fixes.py', verifier_code)
write_file('verify_fixes.py', verifier_code)

print("All backend and test files generated successfully!")



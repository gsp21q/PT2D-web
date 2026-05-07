import React, { useState } from 'react';
import ReactMarkdown from 'react-markdown';
import remarkMath from 'remark-math';
import rehypeKatex from 'rehype-katex';
import 'katex/dist/katex.min.css';
import { Download, FileText, Settings, Type } from 'lucide-react';
import { motion } from 'motion/react';

export default function App() {
  const [text, setText] = useState(`# Main Title

## Subtitle here

This is an example paragraph. You can write your text here and include LaTeX math:

Inline math: Einstein's formula is $E = mc^2$.

Block math:
$$
x = \\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}
$$
`);

  const [direction, setDirection] = useState<'ltr' | 'rtl'>('ltr');
  const [titleColor, setTitleColor] = useState('#1e40af');
  const [subtitleColor, setSubtitleColor] = useState('#0369a1');
  const [bodyColor, setBodyColor] = useState('#333333');
  const [isExporting, setIsExporting] = useState(false);
  const [view, setView] = useState<'editor' | 'preview'>('editor');

  const handleExport = async () => {
    setIsExporting(true);
    try {
      const response = await fetch('/api/convert', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          text,
          direction,
          titleColor,
          subtitleColor,
          bodyColor,
        }),
      });

      if (!response.ok) {
        throw new Error('Export failed');
      }

      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'document.docx';
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(url);
    } catch (err) {
      console.error(err);
      alert('Failed to generate document. Please try again.');
    } finally {
      setIsExporting(false);
    }
  };

  return (
    <div className="flex flex-col h-[100dvh] w-full bg-slate-50 font-sans text-slate-800">
      {/* Top Navigation Bar */}
      <nav className="h-auto md:h-16 py-3 md:py-0 flex flex-col md:flex-row items-center justify-between px-4 md:px-8 bg-white border-b border-slate-200 shrink-0 gap-3 md:gap-0">
        <div className="flex items-center gap-3 w-full md:w-auto justify-between md:justify-start">
          <div className="flex items-center gap-3">
            <div className="w-8 h-8 bg-indigo-600 rounded-lg flex items-center justify-center text-white font-bold tracking-tighter">D</div>
            <span className="text-xl font-semibold tracking-tight">PT<span className="text-indigo-500">2D</span></span>
          </div>
          {/* Mobile view toggle */}
          <div className="flex md:hidden bg-slate-100 p-1 rounded-full">
            <button
              onClick={() => setView('editor')}
              className={`px-3 py-1 rounded-full text-sm font-medium ${view === 'editor' ? 'bg-white shadow-sm' : 'text-slate-500'}`}
            >
              Editor
            </button>
            <button
              onClick={() => setView('preview')}
              className={`px-3 py-1 rounded-full text-sm font-medium ${view === 'preview' ? 'bg-white shadow-sm' : 'text-slate-500'}`}
            >
              Preview
            </button>
          </div>
        </div>
        
        {/* Desktop view toggle */}
        <div className="hidden md:flex items-center gap-6">
          <div className="flex bg-slate-100 p-1 rounded-full">
            <button
              onClick={() => setView('editor')}
              className={`px-4 py-1.5 rounded-full text-sm font-medium ${view === 'editor' ? 'bg-white shadow-sm' : 'text-slate-500'}`}
            >
              Editor
            </button>
            <button
              onClick={() => setView('preview')}
              className={`px-4 py-1.5 rounded-full text-sm font-medium ${view === 'preview' ? 'bg-white shadow-sm' : 'text-slate-500'}`}
            >
              Preview
            </button>
          </div>
        </div>
      </nav>

      {/* Main Content Area */}
      <main className="flex flex-col md:flex-row flex-1 overflow-y-auto md:overflow-hidden">
        
        {/* Left Pane: Input Editor or Preview */}
        <section className="flex-1 flex flex-col p-4 md:p-6 gap-4 border-b md:border-b-0 md:border-r border-slate-200 min-h-[400px] md:min-h-0">
          <div className="flex items-center justify-between">
            <h2 className="text-sm font-bold uppercase tracking-wider text-slate-400">
              {view === 'editor' ? 'Source Input (Markdown/LaTeX)' : 'Live Preview'}
            </h2>
            <div className="flex gap-2">
              <span className="px-2 py-1 bg-green-100 text-green-700 text-xs font-bold rounded">PANDOC v3.1</span>
            </div>
          </div>
          <div className="flex-1 relative">
            {view === 'editor' ? (
              <textarea 
                className="w-full h-full p-6 bg-white border border-slate-200 rounded-xl shadow-inner font-mono text-sm leading-relaxed focus:ring-2 focus:ring-indigo-500 focus:outline-none resize-none"
                value={text}
                onChange={(e) => setText(e.target.value)}
                placeholder="Enter your text or LaTeX here...\n\n# Main Title\n\n$$ E = mc^2 $$\n\nInline math like $a^2 + b^2 = c^2$ is supported too."
                style={{ direction }}
                spellCheck="false"
              />
            ) : (
              <div 
                className="w-full h-full p-6 bg-white border border-slate-200 rounded-xl shadow-inner overflow-y-auto"
                style={{ direction }}
              >
                <ReactMarkdown
                  remarkPlugins={[remarkMath]}
                  rehypePlugins={[rehypeKatex]}
                  components={{
                    h1: ({node, ...props}) => <h1 style={{ color: titleColor }} className="text-4xl font-bold mb-4" {...props} />,
                    h2: ({node, ...props}) => <h2 style={{ color: subtitleColor }} className="text-3xl font-semibold mb-3 mt-6" {...props} />,
                    p: ({node, ...props}) => <p style={{ color: bodyColor }} className="mb-4 text-base leading-relaxed" {...props} />,
                  }}
                >
                  {text}
                </ReactMarkdown>
              </div>
            )}
          </div>
        </section>

        {/* Right Pane: Settings & Controls */}
        <section className="w-full md:w-96 flex flex-col bg-white md:overflow-y-auto p-4 md:p-6 gap-6 md:gap-8 shrink-0">
          
          {/* Document Direction */}
          <div>
            <label className="text-sm font-bold uppercase tracking-wider text-slate-400 block mb-4">Direction & Language</label>
            <div className="grid grid-cols-2 gap-3">
              <button 
                onClick={() => setDirection('ltr')}
                className={`flex flex-col items-center justify-center p-4 border-2 rounded-xl transition-colors ${direction === 'ltr' ? 'border-indigo-500 bg-indigo-50' : 'border-slate-100 hover:border-slate-200'}`}
              >
                <span className={`text-xs font-semibold mb-1 ${direction === 'ltr' ? 'text-indigo-700' : 'text-slate-500'}`}>LTR</span>
                <span className="text-lg">English</span>
              </button>
              <button 
                onClick={() => setDirection('rtl')}
                className={`flex flex-col items-center justify-center p-4 border-2 rounded-xl transition-colors ${direction === 'rtl' ? 'border-indigo-500 bg-indigo-50' : 'border-slate-100 hover:border-slate-200'}`}
              >
                <span className={`text-xs font-semibold mb-1 ${direction === 'rtl' ? 'text-indigo-700' : 'text-slate-500'}`}>RTL</span>
                <span className="text-lg">العربية</span>
              </button>
            </div>
          </div>

          {/* Typography Styles */}
          <div className="space-y-6">
            <label className="text-sm font-bold uppercase tracking-wider text-slate-400 block">Typography Styles</label>
            
            <div className="space-y-4">
              <div className="flex items-center justify-between">
                <span className="text-sm font-medium">Main Title Color</span>
                <input
                  type="color"
                  value={titleColor}
                  onChange={(e) => setTitleColor(e.target.value)}
                  className="w-8 h-8 rounded cursor-pointer border-0 p-0"
                />
              </div>
              <div className="flex items-center justify-between">
                <span className="text-sm font-medium">Subtitle Color</span>
                <input
                  type="color"
                  value={subtitleColor}
                  onChange={(e) => setSubtitleColor(e.target.value)}
                  className="w-8 h-8 rounded cursor-pointer border-0 p-0"
                />
              </div>
              <div className="flex items-center justify-between">
                <span className="text-sm font-medium">Body Text Color</span>
                <input
                  type="color"
                  value={bodyColor}
                  onChange={(e) => setBodyColor(e.target.value)}
                  className="w-8 h-8 rounded cursor-pointer border-0 p-0"
                />
              </div>
            </div>
          </div>

          {/* Export Area */}
          <div className="mt-auto">
            <div className="p-4 bg-slate-50 rounded-xl border border-dashed border-slate-300 mb-4">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 bg-blue-100 text-blue-600 rounded flex items-center justify-center">
                  <svg className="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M7 21h10a2 2 0 002-2V9.414a1 1 0 00-.293-.707l-5.414-5.414A1 1 0 0012.586 3H7a2 2 0 00-2 2v14a2 2 0 002 2z"></path></svg>
                </div>
                <div>
                  <p className="text-xs font-bold text-slate-500">OUTPUT FORMAT</p>
                  <p className="text-sm font-semibold">Microsoft Word (.docx)</p>
                </div>
              </div>
            </div>
            
            <button 
              onClick={handleExport}
              disabled={isExporting}
              className="w-full py-4 bg-indigo-600 text-white rounded-xl font-bold shadow-lg shadow-indigo-200 hover:bg-indigo-700 disabled:opacity-50 disabled:cursor-not-allowed transition-colors flex items-center justify-center gap-2"
            >
              {isExporting ? (
                <span className="animate-spin rounded-full h-5 w-5 border-b-2 border-white"></span>
              ) : (
                <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4"></path></svg>
              )}
              {isExporting ? 'Generating...' : 'Convert & Download'}
            </button>
            <p className="text-[10px] text-center text-slate-400 mt-3">Styles and RTL/LTR flags will be embedded in the docx file.</p>
          </div>
        </section>
      </main>

      {/* Bottom Status Bar */}
      <footer className="h-8 bg-slate-800 text-slate-400 text-[10px] uppercase tracking-widest hidden md:flex items-center px-8 justify-between shrink-0">
        <div className="flex gap-4">
          <span>Words: {text.trim() === '' ? 0 : text.trim().split(/\s+/).length}</span>
          <span>LaTeX Blocks: {Math.floor((text.split('$$').length - 1) / 2)}</span>
          <span>Language: {direction === 'ltr' ? 'EN' : 'AR'}</span>
        </div>
        <div className="flex gap-4">
          <span>Pandoc Kernel: Active</span>
          <span className="text-green-400">● System Ready</span>
        </div>
      </footer>
    </div>
  );
}

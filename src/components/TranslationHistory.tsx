import React from 'react';
import { motion, AnimatePresence } from 'motion/react';
import { saveAs } from 'file-saver';
import { 
  History, 
  FileText, 
  FileSpreadsheet, 
  File as FileIcon, 
  Presentation, 
  Download 
} from 'lucide-react';

interface TranslationHistoryItem {
  name: string;
  date: string;
  blob: Blob;
  type: string;
}

interface TranslationHistoryProps {
  history: TranslationHistoryItem[];
}

export const TranslationHistory: React.FC<TranslationHistoryProps> = React.memo(({ history }) => {
  return (
    <AnimatePresence>
      {history.length > 0 && (
        <motion.div
          initial={{ opacity: 0, y: 20 }}
          animate={{ opacity: 1, y: 0 }}
          className="mt-12 bg-white  rounded-[32px] shadow-[0_8px_30px_rgb(0,0,0,0.04)] border border-white overflow-hidden"
        >
          <div className="p-6 md:p-8 border-b border-slate-100 flex items-center justify-between bg-white/50">
            <div className="flex items-center gap-4">
              <div className="w-10 h-10 rounded-xl bg-[var(--signal-dim)] flex items-center justify-center">
                <History className="w-5 h-5 text-[var(--signal)]" />
              </div>
              <div>
                <h2 className="text-sm font-semibold text-[var(--ink)]">最近翻譯</h2>
                <p className="text-[10px] text-[var(--muted)] uppercase tracking-wider font-medium">Translation History</p>
              </div>
            </div>
          </div>
          <div className="divide-y divide-slate-100">
            {history.map((item, idx) => (
              <div key={idx} className="p-5 md:p-6 flex items-center justify-between hover:bg-white transition-all group">
                <div className="flex items-center gap-5">
                  <div className="w-12 h-12 rounded-2xl bg-[var(--paper)] flex items-center justify-center text-[var(--muted)] group-hover:bg-white group-hover:shadow-sm transition-all border border-slate-100">
                    {item.type === 'docx' && <FileText className="w-6 h-6 text-blue-500" />}
                    {item.type === 'xlsx' && <FileSpreadsheet className="w-6 h-6 text-emerald-500" />}
                    {item.type === 'pdf' && <FileIcon className="w-6 h-6 text-red-500" />}
                    {item.type === 'pptx' && <Presentation className="w-6 h-6 text-orange-500" />}
                    {!['docx', 'xlsx', 'pdf', 'pptx'].includes(item.type) && <FileIcon className="w-6 h-6" />}
                  </div>
                  <div>
                    <h3 className="text-sm font-medium text-slate-700 truncate max-w-[180px] sm:max-w-[400px]">
                      {item.name}
                    </h3>
                    <div className="flex items-center gap-2 mt-1">
                      <span className="text-[10px] text-[var(--muted)] font-light">{item.date}</span>
                      <span className="w-1 h-1 rounded-full bg-slate-300" />
                      <span className="text-[10px] text-[var(--signal)] font-medium uppercase tracking-tighter">Success</span>
                    </div>
                  </div>
                </div>
                <button
                  onClick={() => saveAs(item.blob, item.name)}
                  className="w-10 h-10 rounded-full flex items-center justify-center hover:bg-[var(--signal-dim)] text-[var(--signal)] transition-all group/btn border border-transparent hover:border-[var(--rule)]"
                  title="重新下載"
                >
                  <Download className="w-4 h-4 group-hover/btn:scale-110 transition-transform" />
                </button>
              </div>
            ))}
          </div>
        </motion.div>
      )}
    </AnimatePresence>
  );
});


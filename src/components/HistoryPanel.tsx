import React from 'react';
import { motion, AnimatePresence } from 'motion/react';
import { History, ArrowRight, Clock, File as FileIcon } from 'lucide-react';

interface HistoryPanelProps {
  isOpen: boolean;
  onClose: () => void;
  dbHistory: any[];
}

export const HistoryPanel: React.FC<HistoryPanelProps> = ({ isOpen, onClose, dbHistory }) => {
  return (
    <AnimatePresence>
      {isOpen && (
        <motion.div 
          initial={{ opacity: 0, x: 300 }}
          animate={{ opacity: 1, x: 0 }}
          exit={{ opacity: 0, x: 300 }}
          className="fixed inset-y-0 right-0 w-full sm:w-80 bg-white/80 backdrop-blur-2xl shadow-2xl z-[60] border-l border-white p-6 overflow-y-auto"
        >
          <div className="flex items-center justify-between mb-8">
            <h2 className="text-xl font-bold text-slate-800 flex items-center gap-2">
              <History className="w-5 h-5 text-indigo-600" />
              翻譯紀錄
            </h2>
            <button onClick={onClose} className="text-slate-400 hover:text-slate-800 transition-colors">
              <ArrowRight className="w-6 h-6" />
            </button>
          </div>

          <div className="space-y-4">
            {dbHistory.length === 0 ? (
              <div className="text-center py-12">
                <Clock className="w-12 h-12 text-slate-200 mx-auto mb-4" />
                <p className="text-slate-500 text-sm">尚無翻譯紀錄</p>
              </div>
            ) : (
              dbHistory.map((item) => (
                <div key={item.id} className="p-4 rounded-xl border border-slate-200 bg-white/50 hover:bg-white transition-all group shadow-sm">
                  <div className="flex items-start gap-3">
                    <div className="p-2 bg-indigo-50 rounded-lg border border-indigo-100">
                      <FileIcon className="w-4 h-4 text-indigo-600" />
                    </div>
                    <div className="flex-1 min-w-0">
                      <p className="text-sm font-semibold text-slate-700 truncate">{item.fileName}</p>
                      <p className="text-[10px] text-slate-500 mt-0.5">
                        {item.timestamp?.toDate().toLocaleString()}
                      </p>
                      <div className="flex flex-wrap gap-1 mt-2">
                        {item.targetLanguages?.map((l: string) => (
                          <span key={l} className="px-1.5 py-0.5 bg-indigo-100 text-indigo-700 text-[9px] font-bold rounded uppercase">
                            {l}
                          </span>
                        ))}
                      </div>
                    </div>
                  </div>
                </div>
              ))
            )}
          </div>
        </motion.div>
      )}
    </AnimatePresence>
  );
};


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
          className="fixed inset-y-0 right-0 w-full sm:w-80 bg-white shadow-2xl z-[60] border-l border-gray-100 p-6 overflow-y-auto"
        >
          <div className="flex items-center justify-between mb-8">
            <h2 className="text-xl font-bold text-gray-900 flex items-center gap-2">
              <History className="w-5 h-5 text-emerald-600" />
              翻譯紀錄
            </h2>
            <button onClick={onClose} className="text-gray-400 hover:text-gray-600">
              <ArrowRight className="w-6 h-6" />
            </button>
          </div>

          <div className="space-y-4">
            {dbHistory.length === 0 ? (
              <div className="text-center py-12">
                <Clock className="w-12 h-12 text-gray-100 mx-auto mb-4" />
                <p className="text-gray-400 text-sm">尚無翻譯紀錄</p>
              </div>
            ) : (
              dbHistory.map((item) => (
                <div key={item.id} className="p-4 rounded-xl border border-gray-50 bg-gray-50/50 hover:bg-white hover:shadow-sm transition-all group">
                  <div className="flex items-start gap-3">
                    <div className="p-2 bg-white rounded-lg border border-gray-100">
                      <FileIcon className="w-4 h-4 text-emerald-600" />
                    </div>
                    <div className="flex-1 min-w-0">
                      <p className="text-sm font-semibold text-gray-900 truncate">{item.fileName}</p>
                      <p className="text-[10px] text-gray-400 mt-0.5">
                        {item.timestamp?.toDate().toLocaleString()}
                      </p>
                      <div className="flex flex-wrap gap-1 mt-2">
                        {item.targetLanguages?.map((l: string) => (
                          <span key={l} className="px-1.5 py-0.5 bg-emerald-50 text-emerald-600 text-[9px] font-bold rounded uppercase">
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

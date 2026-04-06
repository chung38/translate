import React from 'react';
import { motion, AnimatePresence } from 'motion/react';

interface TranslationProgressProps {
  status: string;
  statusMessage: string;
  progress: number;
  files?: File[];
  fileProgress?: Record<string, number>;
}

export const TranslationProgress: React.FC<TranslationProgressProps> = React.memo(({
  status,
  statusMessage,
  progress,
  files = [],
  fileProgress = {}
}) => {
  return (
    <AnimatePresence>
      {(status !== 'idle' && status !== 'error') && (
        <motion.div 
          initial={{ opacity: 0, height: 0 }}
          animate={{ opacity: 1, height: 'auto' }}
          exit={{ opacity: 0, height: 0 }}
          className="mt-8 pt-8 border-t border-slate-200"
        >
          <div className="flex justify-between items-center mb-2">
            <span className="text-sm font-medium text-slate-600">
              {statusMessage}
            </span>
            {(status === 'processing' || status === 'translating' || status === 'generating') && (
              <span className="text-xs font-mono text-indigo-600">{progress}%</span>
            )}
          </div>
          <div className="w-full h-1.5 bg-slate-200 rounded-full overflow-hidden mb-4">
            <motion.div 
              initial={{ width: 0 }}
              animate={{ width: `${progress}%` }}
              className="h-full bg-indigo-500 shadow-[0_0_10px_rgba(99,102,241,0.5)]"
            />
          </div>

          {files.length > 1 && (
            <div className="space-y-3 mt-4">
              {files.map((file, index) => {
                let currentProgress = 0;
                if (fileProgress[file.name] !== undefined) {
                  currentProgress = fileProgress[file.name];
                } else {
                  const hasSubsequentProgress = files.slice(index + 1).some(f => fileProgress[f.name] !== undefined);
                  if (hasSubsequentProgress) {
                    currentProgress = 100;
                  } else if (status === 'completed') {
                    currentProgress = 100;
                  }
                }

                if (status === 'completed') {
                  currentProgress = 100;
                }

                return (
                  <div key={file.name} className="flex flex-col gap-1">
                    <div className="flex justify-between items-center">
                      <span className="text-xs text-slate-500 truncate pr-4" title={file.name}>
                        {file.name}
                      </span>
                      <span className="text-xs font-mono text-indigo-500 shrink-0">
                        {currentProgress}%
                      </span>
                    </div>
                    <div className="w-full h-1 bg-slate-200 rounded-full overflow-hidden">
                      <motion.div 
                        initial={{ width: 0 }}
                        animate={{ width: `${currentProgress}%` }}
                        className="h-full bg-indigo-400 shadow-[0_0_8px_rgba(99,102,241,0.5)]"
                      />
                    </div>
                  </div>
                );
              })}
            </div>
          )}
        </motion.div>
      )}
    </AnimatePresence>
  );
});


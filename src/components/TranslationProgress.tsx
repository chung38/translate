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
          className="mt-8 pt-8 border-top border-gray-100"
        >
          <div className="flex justify-between items-center mb-2">
            <span className="text-sm font-medium text-gray-600">
              {statusMessage}
            </span>
            {(status === 'processing' || status === 'translating' || status === 'generating') && (
              <span className="text-xs font-mono text-emerald-600">{progress}%</span>
            )}
          </div>
          <div className="w-full h-1.5 bg-gray-100 rounded-full overflow-hidden mb-4">
            <motion.div 
              initial={{ width: 0 }}
              animate={{ width: `${progress}%` }}
              className="h-full bg-emerald-500"
            />
          </div>

          {files.length > 1 && (
            <div className="space-y-3 mt-4">
              {files.map((file, index) => {
                // Determine file status based on overall progress and fileProgress
                // Since files are processed sequentially, if a file has fileProgress, it's either processing or completed.
                // If its progress is 100, it's completed.
                // If it doesn't have fileProgress yet, it's pending (0%).
                // Wait, if a file is completed, does fileProgress keep it at 100?
                // Let's assume fileProgress[file.name] is updated to 100 when done, or we can infer it based on index.
                // Actually, the current file being processed has fileProgress[file.name].
                // Previous files are 100%. Future files are 0%.
                // We can find the current processing index by looking at the first file that has progress < 100, or just use the fileProgress map.
                
                let currentProgress = 0;
                if (fileProgress[file.name] !== undefined) {
                  currentProgress = fileProgress[file.name];
                } else {
                  // If it's not in fileProgress, check if any subsequent file has progress.
                  // Since it's sequential, if a file later in the list has progress, this one must be 100%.
                  // If no subsequent file has progress, this one is 0%.
                  const hasSubsequentProgress = files.slice(index + 1).some(f => fileProgress[f.name] !== undefined);
                  if (hasSubsequentProgress) {
                    currentProgress = 100;
                  } else if (status === 'completed') {
                    currentProgress = 100;
                  }
                }

                // If status is completed, all files are 100%
                if (status === 'completed') {
                  currentProgress = 100;
                }

                return (
                  <div key={file.name} className="flex flex-col gap-1">
                    <div className="flex justify-between items-center">
                      <span className="text-xs text-gray-500 truncate pr-4" title={file.name}>
                        {file.name}
                      </span>
                      <span className="text-xs font-mono text-emerald-600 shrink-0">
                        {currentProgress}%
                      </span>
                    </div>
                    <div className="w-full h-1 bg-gray-100 rounded-full overflow-hidden">
                      <motion.div 
                        initial={{ width: 0 }}
                        animate={{ width: `${currentProgress}%` }}
                        className="h-full bg-emerald-400"
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

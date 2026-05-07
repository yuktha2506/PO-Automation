import React from 'react';

interface ProgressBarProps {
  current: number;
  total: number;
}

export const ProgressBar: React.FC<ProgressBarProps> = ({ current, total }) => {
  const percentage = total === 0 ? 0 : Math.round((current / total) * 100);

  return (
    <div className="w-full space-y-3">
      <div className="flex justify-between text-sm font-semibold text-gray-700 dark:text-gray-300">
        <span>Processing files...</span>
        <span>{current}/{total} <span className="text-gray-400 dark:text-gray-500 font-normal">({percentage}%)</span></span>
      </div>
      <div className="w-full bg-gray-200 dark:bg-slate-700 rounded-full h-3 overflow-hidden shadow-inner">
        <div 
          className="bg-gradient-to-r from-blue-500 to-indigo-600 h-3 rounded-full transition-all duration-500 ease-out relative" 
          style={{ width: `${percentage}%` }}
        >
          {/* Shimmer effect */}
          <div className="absolute inset-0 bg-white/30 w-full h-full animate-[shimmer_2s_infinite] -skew-x-12 transform -translate-x-full" style={{ backgroundImage: 'linear-gradient(90deg, transparent, rgba(255,255,255,0.5), transparent)' }}></div>
        </div>
      </div>
    </div>
  );
};
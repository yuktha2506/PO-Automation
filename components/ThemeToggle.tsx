import React from 'react';
import { Moon, Sun } from 'lucide-react';
import { useTheme } from './ThemeContext';

export const ThemeToggle = () => {
  const { theme, toggleTheme } = useTheme();

  return (
    <button
      onClick={toggleTheme}
      className="p-2 rounded-full transition-all duration-300 hover:bg-gray-100 dark:hover:bg-slate-700 focus:outline-none focus:ring-2 focus:ring-blue-500 text-gray-600 dark:text-slate-300"
      aria-label="Toggle Dark Mode"
    >
      <div className="relative w-6 h-6">
        <Sun 
          className={`absolute inset-0 w-6 h-6 transition-transform duration-500 rotate-0 scale-100 dark:-rotate-90 dark:scale-0 text-yellow-500`} 
        />
        <Moon 
          className={`absolute inset-0 w-6 h-6 transition-transform duration-500 rotate-90 scale-0 dark:rotate-0 dark:scale-100 text-slate-100`} 
        />
      </div>
    </button>
  );
};
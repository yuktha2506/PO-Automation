import React from 'react';
import { Dashboard } from './components/Dashboard';
import { ThemeProvider } from './components/ThemeContext';
import { ThemeToggle } from './components/ThemeToggle';

function AppContent() {
  return (
    <div className="min-h-screen bg-gray-50 dark:bg-slate-900 transition-colors duration-500">
      <nav className="bg-white/70 dark:bg-slate-900/70 backdrop-blur-xl border-b border-gray-200/50 dark:border-slate-800/50 sticky top-0 z-50 transition-all duration-300">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
          <div className="flex justify-between h-18 items-center py-3">
            <div className="flex items-center space-x-3 group cursor-pointer">
              <div className="bg-gradient-to-tr from-blue-600 to-indigo-600 rounded-xl p-2 shadow-lg shadow-blue-500/30 group-hover:shadow-blue-500/50 transition-all duration-300">
                 <svg className="w-6 h-6 text-white" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2.5} d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" /></svg>
              </div>
              <span className="text-gray-900 dark:text-white font-extrabold text-2xl tracking-tight bg-clip-text text-transparent bg-gradient-to-r from-gray-900 to-gray-700 dark:from-white dark:to-gray-300">AutoPO</span>
            </div>
            <div className="flex items-center space-x-4">
               <div className="hidden md:block h-6 w-px bg-gray-200 dark:bg-slate-700"></div>
               <ThemeToggle />
            </div>
          </div>
        </div>
      </nav>
      <main className="transition-all duration-500">
        <Dashboard />
      </main>
    </div>
  );
}

function App() {
  return (
    <ThemeProvider>
      <AppContent />
    </ThemeProvider>
  );
}

export default App;
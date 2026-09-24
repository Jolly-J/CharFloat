import React, { useState } from 'react';
import { Navbar } from './components/Navbar';
import { Hero } from './components/Hero';
import { McpPluginSection } from './components/McpPluginSection';
import { DeliverySection } from './components/DeliverySection';
import { ComparisonTable } from './components/ComparisonTable';
import { UseCasesMarquee } from './components/UseCasesMarquee';
import { AgentGrid } from './components/AgentGrid';
import { Footer } from './components/Footer';
import { QuickStartModal } from './components/QuickStartModal';
import { DownloadModal } from './components/DownloadModal';
import { getDetectedOS } from './utils/os';

export const App: React.FC = () => {
  const [isQuickStartOpen, setIsQuickStartOpen] = useState(false);
  const [isDownloadOpen, setIsDownloadOpen] = useState(false);
  const [downloadTab, setDownloadTab] = useState<'mac' | 'win' | 'cli'>(() => getDetectedOS());

  const handleOpenDownload = (tab?: 'mac' | 'win' | 'cli') => {
    const targetTab = tab || getDetectedOS();
    setDownloadTab(targetTab);
    setIsDownloadOpen(true);
  };

  return (
    <div 
      className="min-h-screen bg-[#f8fafc] text-slate-800 flex flex-col font-sans selection:bg-blue-500/20 selection:text-blue-700"
      style={{
        backgroundImage: 'linear-gradient(180deg, #cae0f9 0%, #e2f0fd 320px, #f2f7fd 640px, #f8fafc 880px)',
        backgroundRepeat: 'no-repeat'
      }}
    >
      {/* 1. Floating Pill Navigation Bar (Doubao Work 1:1) */}
      <Navbar 
        onOpenQuickStart={() => setIsQuickStartOpen(true)} 
        onOpenDownload={handleOpenDownload}
      />

      {/* Main Content Sections */}
      <main className="flex-1">
        {/* 2. Hero Section with Integrated Showcase Banner */}
        <Hero 
          onOpenQuickStart={() => setIsQuickStartOpen(true)} 
          onOpenDownload={handleOpenDownload}
        />

        {/* 2.5 Free MCP Plugin & Client Showcase (Left Image + Right Interactive Tabs) */}
        <McpPluginSection />

        {/* 3. Deep Delivery Section (Vertical Tabs + Live Large Preview) */}
        <DeliverySection />

        {/* 4. Two-track demo: regenerated files vs edits in the current workspace */}
        <ComparisonTable />

        {/* 5. Model & Ecosystem Support Grid */}
        <AgentGrid onOpenQuickStart={() => setIsQuickStartOpen(true)} />

        {/* 6. Real Prompts & Use Cases Marquee (Double Row Horizontal Auto-Scroll) */}
        <UseCasesMarquee />
      </main>

      {/* 7. Footer Closing CTA & Legal Bar */}
      <Footer 
        onOpenQuickStart={() => setIsQuickStartOpen(true)} 
        onOpenDownload={handleOpenDownload}
      />

      {/* Download Center Modal */}
      <DownloadModal
        isOpen={isDownloadOpen}
        initialTab={downloadTab}
        onClose={() => setIsDownloadOpen(false)}
        onOpenQuickStart={() => {
          setIsDownloadOpen(false);
          setIsQuickStartOpen(true);
        }}
      />

      {/* Quick Start Modal */}
      <QuickStartModal 
        isOpen={isQuickStartOpen} 
        onClose={() => setIsQuickStartOpen(false)} 
      />
    </div>
  );
};

export default App;

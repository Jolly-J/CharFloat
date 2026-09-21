import React, { useState } from 'react';
import { Hero } from './components/Hero';
import { LiveWorkspace } from './components/LiveWorkspace';
import { ArchitectureFlow } from './components/ArchitectureFlow';
import { ComparisonTable } from './components/ComparisonTable';
import { ValueProps } from './components/ValueProps';
import { RealCaseTimeline } from './components/RealCaseTimeline';
import { AgentGrid } from './components/AgentGrid';
import { Footer } from './components/Footer';
import { QuickStartModal } from './components/QuickStartModal';

export const App: React.FC = () => {
  const [isQuickStartOpen, setIsQuickStartOpen] = useState(false);

  return (
    <div className="min-h-screen bg-[#08080a] text-zinc-100 flex flex-col font-sans selection:bg-blue-500/30 selection:text-blue-200">
      {/* Main Content Sections */}
      <main className="flex-1">
        <Hero />
        <LiveWorkspace />
        <ArchitectureFlow />
        <ComparisonTable />
        <ValueProps />
        <RealCaseTimeline />
        <AgentGrid onOpenQuickStart={() => setIsQuickStartOpen(true)} />
      </main>

      {/* Footer */}
      <Footer onOpenQuickStart={() => setIsQuickStartOpen(true)} />

      {/* Quick Start Modal */}
      <QuickStartModal 
        isOpen={isQuickStartOpen} 
        onClose={() => setIsQuickStartOpen(false)} 
      />
    </div>
  );
};

export default App;

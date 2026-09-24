import React from 'react';

export const AppleLogo: React.FC<{ className?: string }> = ({ className = 'w-4 h-4' }) => (
  <img 
    src="/icon-apple.png" 
    alt="Apple macOS" 
    className={`${className} object-contain shrink-0 select-none`}
    loading="lazy"
  />
);

export const WindowsLogo: React.FC<{ className?: string }> = ({ className = 'w-4 h-4' }) => (
  <img 
    src="/icon-windows.png" 
    alt="Microsoft Windows" 
    className={`${className} object-contain shrink-0 select-none`}
    loading="lazy"
  />
);

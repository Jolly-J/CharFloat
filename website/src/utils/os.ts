/**
 * Automatically detects the user's operating system platform.
 * Supports Windows, macOS, and falls back to Windows on general desktop.
 */
export function getDetectedOS(): 'mac' | 'win' | 'cli' {
  if (typeof window === 'undefined') return 'win';

  // 1. Modern Client Hints API (navigator.userAgentData.platform)
  const userAgentData = (window.navigator as any)?.userAgentData;
  if (userAgentData?.platform) {
    const p = String(userAgentData.platform).toLowerCase();
    if (p.includes('win')) return 'win';
    if (p.includes('mac') || p.includes('darwin')) return 'mac';
  }

  // 2. Navigator User Agent & Platform Strings
  const ua = window.navigator.userAgent.toLowerCase();
  const legacyPlatform = window.navigator.platform ? window.navigator.platform.toLowerCase() : '';

  if (ua.includes('windows') || ua.includes('win32') || ua.includes('win64') || legacyPlatform.includes('win')) {
    return 'win';
  }

  if (ua.includes('macintosh') || ua.includes('mac os') || legacyPlatform.includes('mac') || ua.includes('darwin')) {
    return 'mac';
  }

  return 'win';
}

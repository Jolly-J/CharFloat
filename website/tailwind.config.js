import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

/** @type {import('tailwindcss').Config} */
export default {
  content: [
    path.resolve(__dirname, "./index.html"),
    path.resolve(__dirname, "./src/**/*.{js,ts,jsx,tsx}"),
  ],
  darkMode: "class",
  theme: {
    extend: {
      colors: {
        bridge: {
          dark: "#08080a",
          card: "#121216",
          elevated: "#1a1a22",
          border: "rgba(255, 255, 255, 0.08)",
          accent: "#3b82f6",
          emerald: "#10b981",
          amber: "#f59e0b",
          purple: "#8b5cf6",
        }
      },
      animation: {
        'glow-slow': 'glow 8s ease-in-out infinite alternate',
        'pulse-subtle': 'subtlePulse 3s cubic-bezier(0.4, 0, 0.6, 1) infinite',
      },
      keyframes: {
        glow: {
          '0%': { opacity: '0.4', transform: 'scale(0.98)' },
          '100%': { opacity: '0.8', transform: 'scale(1.02)' },
        },
        subtlePulse: {
          '0%, 100%': { opacity: '1' },
          '50%': { opacity: '0.5' },
        }
      }
    },
  },
  plugins: [],
};

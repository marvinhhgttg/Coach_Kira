import type { Config } from 'tailwindcss';

export default {
  darkMode: ['class'],
  content: ['./index.html', './src/**/*.{ts,tsx}'],
  theme: {
    extend: {
      fontFamily: {
        sans: ['Inter', 'ui-sans-serif', 'system-ui', '-apple-system', 'sans-serif'],
        mono: ['JetBrains Mono', 'ui-monospace', 'SFMono-Regular', 'monospace'],
      },
      colors: {
        bg: {
          DEFAULT: '#0a0d12',
          raised: '#0f1319',
          panel: '#131820',
          subtle: '#1a212b',
        },
        border: {
          DEFAULT: '#1f2731',
          strong: '#2a3340',
        },
        ink: {
          DEFAULT: '#e6ebf2',
          muted: '#8a96a8',
          dim: '#5b6678',
        },
        accent: {
          DEFAULT: '#7dd3fc',
          strong: '#38bdf8',
        },
        ampel: {
          gruen: '#22c55e',
          gelb: '#eab308',
          orange: '#f97316',
          rot: '#ef4444',
          grau: '#6b7280',
          blau: '#3b82f6',
          lila: '#a855f7',
        },
      },
      borderRadius: {
        DEFAULT: '6px',
        md: '8px',
        lg: '10px',
      },
      fontSize: {
        '2xs': ['0.6875rem', { lineHeight: '1rem' }],
      },
    },
  },
  plugins: [],
} satisfies Config;

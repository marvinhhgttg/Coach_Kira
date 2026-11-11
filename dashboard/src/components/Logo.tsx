export function Logo({ size = 22 }: { size?: number }) {
  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 32 32"
      fill="none"
      aria-label="Coach Kira"
      className="shrink-0"
    >
      <rect x="0.5" y="0.5" width="31" height="31" rx="6" stroke="currentColor" strokeOpacity="0.25" />
      <path
        d="M8 22 L8 10 M8 16 L18 10 M8 16 L18 22 M22 10 L22 22"
        stroke="currentColor"
        strokeWidth="2"
        strokeLinecap="square"
      />
    </svg>
  );
}

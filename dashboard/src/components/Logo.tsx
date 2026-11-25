// 25.11.2026: App-Symbol (Bogen mit Berg und Herzschlag) statt "KI" – gleich wie iPhone-/Mac-App
import logo from '../assets/kira-logo.png';

export function Logo({ size = 28 }: { size?: number }) {
  return <img src={logo} width={size} height={size} alt="Coach Kira" className="shrink-0 rounded-[22%]" draggable={false} />;
}

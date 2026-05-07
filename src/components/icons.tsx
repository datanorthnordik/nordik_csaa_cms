import { Menu, Notifications, Help, ConfirmationNumber, Dashboard, Description, Event, Email, Newspaper, AccountCircle, Folder, Image, Settings, Logout } from '@mui/icons-material';
import type { CSSProperties } from 'react';

type IconProps = {
  size?: number
  className?: string
}

const getIconStyle = (size: number): CSSProperties => ({
  width: size,
  height: size,
  fontSize: size,
});

export function MenuIcon({ size = 22, className }: IconProps) {
  return <Menu style={getIconStyle(size)} className={className} />;
}

export function BellIcon({ size = 22, className }: IconProps) {
  return <Notifications style={getIconStyle(size)} className={className} />;
}

export function HelpIcon({ size = 22, className }: IconProps) {
  return <Help style={getIconStyle(size)} className={className} />;
}

export function TicketIcon({ size = 16, className }: IconProps) {
  return <ConfirmationNumber style={getIconStyle(size)} className={className} />;
}

export function DashboardIcon({ size = 18, className }: IconProps) {
  return <Dashboard style={getIconStyle(size)} className={className} />;
}

export function PagesIcon({ size = 18, className }: IconProps) {
  return <Description style={getIconStyle(size)} className={className} />;
}

export function EventsIcon({ size = 18, className }: IconProps) {
  return <Event style={getIconStyle(size)} className={className} />;
}

export function NewslettersIcon({ size = 18, className }: IconProps) {
  return <Email style={getIconStyle(size)} className={className} />;
}

export function PressIcon({ size = 18, className }: IconProps) {
  return <Newspaper style={getIconStyle(size)} className={className} />;
}

export function MemorialIcon({ size = 18, className }: IconProps) {
  return <AccountCircle style={getIconStyle(size)} className={className} />;
}

export function ResourcesIcon({ size = 18, className }: IconProps) {
  return <Folder style={getIconStyle(size)} className={className} />;
}

export function MediaIcon({ size = 18, className }: IconProps) {
  return <Image style={getIconStyle(size)} className={className} />;
}

export function SettingsIcon({ size = 18, className }: IconProps) {
  return <Settings style={getIconStyle(size)} className={className} />;
}

export function LogoutIcon({ size = 18, className }: IconProps) {
  return <Logout style={getIconStyle(size)} className={className} />;
}

import {
  Menu,
  Notifications,
  Help,
  ConfirmationNumber,
  Dashboard,
  Description,
  Event,
  Email,
  Newspaper,
  AccountCircle,
  Folder,
  Image,
  Settings,
  Logout,
  Add,
  Search,
  Close,
  CloudUpload,
  Download,
  Delete,
  Edit,
  Link as LinkIconMui,
  CheckCircle,
  Tune,
  CollectionsBookmark,
  AddPhotoAlternate,
  ChevronLeft,
  ChevronRight,
} from '@mui/icons-material';
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

export function AddIcon({ size = 18, className }: IconProps) {
  return <Add style={getIconStyle(size)} className={className} />;
}

export function SearchIcon({ size = 18, className }: IconProps) {
  return <Search style={getIconStyle(size)} className={className} />;
}

export function CloseIcon({ size = 18, className }: IconProps) {
  return <Close style={getIconStyle(size)} className={className} />;
}

export function CloudUploadIcon({ size = 22, className }: IconProps) {
  return <CloudUpload style={getIconStyle(size)} className={className} />;
}

export function DownloadIcon({ size = 18, className }: IconProps) {
  return <Download style={getIconStyle(size)} className={className} />;
}

export function DeleteIcon({ size = 18, className }: IconProps) {
  return <Delete style={getIconStyle(size)} className={className} />;
}

export function EditIcon({ size = 14, className }: IconProps) {
  return <Edit style={getIconStyle(size)} className={className} />;
}

export function LinkIcon({ size = 14, className }: IconProps) {
  return <LinkIconMui style={getIconStyle(size)} className={className} />;
}

export function CheckCircleIcon({ size = 14, className }: IconProps) {
  return <CheckCircle style={getIconStyle(size)} className={className} />;
}

export function SpecsIcon({ size = 14, className }: IconProps) {
  return <Tune style={getIconStyle(size)} className={className} />;
}

export function GalleryIcon({ size = 14, className }: IconProps) {
  return <CollectionsBookmark style={getIconStyle(size)} className={className} />;
}

export function AddPhotoIcon({ size = 22, className }: IconProps) {
  return <AddPhotoAlternate style={getIconStyle(size)} className={className} />;
}

export function ChevronLeftIcon({ size = 18, className }: IconProps) {
  return <ChevronLeft style={getIconStyle(size)} className={className} />;
}

export function ChevronRightIcon({ size = 18, className }: IconProps) {
  return <ChevronRight style={getIconStyle(size)} className={className} />;
}

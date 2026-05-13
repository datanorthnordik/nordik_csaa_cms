import { render } from '@testing-library/react'
import { describe, expect, it } from 'vitest'
import {
  AddIcon,
  AddPhotoIcon,
  BellIcon,
  CheckCircleIcon,
  ChevronLeftIcon,
  ChevronRightIcon,
  CloudUploadIcon,
  CloseIcon,
  DashboardIcon,
  DeleteIcon,
  DownloadIcon,
  EditIcon,
  EventsIcon,
  GalleryIcon,
  HelpIcon,
  LinkIcon,
  LogoutIcon,
  MediaIcon,
  MemorialIcon,
  MenuIcon,
  MenusIcon,
  NewslettersIcon,
  PagesIcon,
  PressIcon,
  ResourcesIcon,
  SearchIcon,
  SettingsIcon,
  SpecsIcon,
  TicketIcon,
} from './icons'

describe('icons', () => {
  it('renders each shared CMS icon component', () => {
    const { container } = render(
      <>
        <MenuIcon />
        <BellIcon />
        <HelpIcon />
        <TicketIcon />
        <DashboardIcon />
        <PagesIcon />
        <MenusIcon />
        <EventsIcon />
        <NewslettersIcon />
        <PressIcon />
        <MemorialIcon />
        <ResourcesIcon />
        <MediaIcon />
        <SettingsIcon />
        <LogoutIcon />
        <AddIcon />
        <SearchIcon />
        <CloseIcon />
        <CloudUploadIcon />
        <DownloadIcon />
        <DeleteIcon />
        <EditIcon />
        <LinkIcon />
        <CheckCircleIcon />
        <SpecsIcon />
        <GalleryIcon />
        <AddPhotoIcon />
        <ChevronLeftIcon />
        <ChevronRightIcon />
      </>,
    )

    expect(container.querySelectorAll('svg')).toHaveLength(29)
  })
})

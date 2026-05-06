import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import { Provider } from 'react-redux'
import { Toaster } from 'react-hot-toast'
import './index.css'
import './i18n'
import App from './App.tsx'
import { store } from './store/store'

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <Provider store={store}>
      <App />
      <Toaster
        position="top-right"
        toastOptions={{
          duration: 4000,
          style: {
            background: '#1f1f23',
            color: '#fff',
            borderRadius: '10px',
          },
          error: {
            style: {
              background: '#7a1f24',
            },
          },
        }}
      />
    </Provider>
  </StrictMode>,
)

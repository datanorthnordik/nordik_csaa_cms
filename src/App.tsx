import { useState } from 'react'
import { LoginPage } from './pages/LoginPage'
import { SignupPage } from './pages/SignupPage'
import './App.css'

type Route = 'login' | 'signup'

function App() {
  const [route, setRoute] = useState<Route>('login')

  if (route === 'signup') {
    return (
      <SignupPage
        onSubmit={(values) => {
          console.log('signup', values)
        }}
        onSignIn={() => setRoute('login')}
      />
    )
  }

  return (
    <LoginPage
      onSubmit={(values) => {
        console.log('login', values)
      }}
      onCreateAccount={() => setRoute('signup')}
      onForgotPassword={() => console.log('forgot password')}
    />
  )
}

export default App

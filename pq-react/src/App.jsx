import { useState } from 'react';
import { StoreProvider } from './store/StoreContext.jsx';
import { useAutoCalcFields } from './store/useAutoCalcFields.js';
import Topbar from './components/Topbar.jsx';
import PQScreen from './screens/PQScreen.jsx';
import PresentationScreen from './screens/PresentationScreen.jsx';
import PlanScreen from './screens/PlanScreen.jsx';
import NotesDrawer from './components/NotesDrawer.jsx';
import DataConsole from './components/DataConsole.jsx';

const SCREENS = {
  PQ: 'pq',
  PRESENTATION: 'presentation',
  PLAN: 'plan',
};

function AppShell() {
  // Runs derived-field calculations on every store change. Lives inside the
  // provider so it can use useStore via useAutoCalcFields.
  useAutoCalcFields();

  const [screen, setScreen] = useState(SCREENS.PQ);
  const [consoleOpen, setConsoleOpen] = useState(false);
  const [drawerOpen, setDrawerOpen] = useState(false);

  return (
    <>
      {screen === SCREENS.PQ && (
        <div className="app-shell">
          <Topbar
            onPresentation={() => setScreen(SCREENS.PRESENTATION)}
            onPlan={() => setScreen(SCREENS.PLAN)}
            onOpenConsole={() => setConsoleOpen(true)}
          />
          <main className="content">
            <PQScreen />
          </main>
        </div>
      )}

      {screen === SCREENS.PRESENTATION && (
        <PresentationScreen onBack={() => setScreen(SCREENS.PQ)} />
      )}

      {screen === SCREENS.PLAN && (
        <PlanScreen onBack={() => setScreen(SCREENS.PQ)} />
      )}

      <NotesDrawer open={drawerOpen} onToggle={() => setDrawerOpen(o => !o)} />
      <DataConsole open={consoleOpen} onClose={() => setConsoleOpen(false)} />
    </>
  );
}

export default function App() {
  return (
    <StoreProvider>
      <AppShell />
    </StoreProvider>
  );
}

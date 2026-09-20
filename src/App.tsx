import { useState } from 'react';
import './styles/theme.css';
import './styles/components.css';
import Sidebar from './components/Sidebar';
import Header from './components/Header';
import Dashboard from './pages/Dashboard';
import Spreadsheet from './pages/Spreadsheet';
import FileEnhancer from './pages/FileEnhancer';
import ChartsPivot from './pages/ChartsPivot';
import AutoReporting from './pages/AutoReporting';
import Settings from './pages/Settings';
import type { NavItem } from './types/navigation';

const navigation: NavItem[] = [
  { id: 'dashboard', label: 'Dashboard', icon: 'dashboard' },
  { id: 'spreadsheet', label: 'Spreadsheet', icon: 'table' },
  { id: 'file-enhancer', label: 'File Enhancer', icon: 'file-cog' },
  { id: 'charts-pivot', label: 'Charts & Pivot', icon: 'chart-bar' },
  { id: 'auto-reporting', label: 'Auto Reporting', icon: 'file-text' },
  { id: 'settings', label: 'Settings', icon: 'settings' },
];

const pageComponents: Record<string, React.ComponentType> = {
  dashboard: Dashboard,
  spreadsheet: Spreadsheet,
  'file-enhancer': FileEnhancer,
  'charts-pivot': ChartsPivot,
  'auto-reporting': AutoReporting,
  settings: Settings,
};

function App() {
  const [currentPage, setCurrentPage] = useState('dashboard');
  const [sidebarCollapsed, setSidebarCollapsed] = useState(false);

  const CurrentPageComponent = pageComponents[currentPage] || Dashboard;

  return (
    <div className="app">
      <Sidebar
        items={navigation}
        activeItem={currentPage}
        onNavigate={setCurrentPage}
        collapsed={sidebarCollapsed}
        onToggleCollapse={() => setSidebarCollapsed(!sidebarCollapsed)}
      />
      <div className="app-main">
        <Header onMenuClick={() => setSidebarCollapsed(!sidebarCollapsed)} />
        <main className="app-content">
          <CurrentPageComponent />
        </main>
      </div>
    </div>
  );
}

export default App;
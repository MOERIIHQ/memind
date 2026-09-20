import { useState } from 'react';
import type { FormEvent } from 'react';
import Button from '../components/Button';
import Card, { CardHeader, CardBody } from '../components/Card';
import Icon from '../components/Icon';
import type { FileItem, QuickAction } from '../types';
import './Dashboard.css';

const mockFiles: FileItem[] = [
  { id: '1', name: 'Q3_Budget_Forecast.xlsx', size: 245760, modified: new Date('2024-12-15'), type: 'xlsx', isMock: true },
  { id: '2', name: 'Employee_Roster_2024.csv', size: 89234, modified: new Date('2024-12-10'), type: 'csv', isMock: true },
  { id: '3', name: 'Sales_Data_November.ods', size: 512000, modified: new Date('2024-12-08'), type: 'ods', isMock: true },
  { id: '4', name: 'Project_Timeline.xlsx', size: 178945, modified: new Date('2024-12-05'), type: 'xlsx', isMock: true },
];

const quickActions: QuickAction[] = [
  {
    id: 'clean',
    label: 'Clean Spreadsheet',
    description: 'Remove duplicates, fix formatting, standardize data',
    icon: 'filter',
  },
  {
    id: 'formula',
    label: 'Generate Formula',
    description: 'Create complex formulas from natural language',
    icon: 'columns',
  },
  {
    id: 'chart',
    label: 'Create Chart',
    description: 'Visualize data with automatic chart suggestions',
    icon: 'chart-bar',
  },
  {
    id: 'report',
    label: 'Generate Report',
    description: 'Transform analysis into formatted reports',
    icon: 'file-text',
  },
];

const Dashboard = () => {
  const [commandInput, setCommandInput] = useState('');

  const handleCommandSubmit = (e: FormEvent) => {
    e.preventDefault();
    if (commandInput.trim()) {
      // Command processing will be implemented in future iterations
      console.log('Command submitted:', commandInput);
    }
  };

  const handleQuickAction = (action: QuickAction) => {
    console.log('Quick action:', action.id);
    // Navigation to respective pages will be implemented
  };

  const formatFileSize = (bytes: number): string => {
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  };

  const formatDate = (date: Date): string => {
    const now = new Date();
    const diff = now.getTime() - date.getTime();
    const days = Math.floor(diff / (1000 * 60 * 60 * 24));
    if (days === 0) return 'Today';
    if (days === 1) return 'Yesterday';
    if (days < 7) return `${days} days ago`;
    return date.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
  };

  const getFileIcon = (type: FileItem['type']) => {
    switch (type) {
      case 'xlsx':
        return 'file-spreadsheet';
      case 'csv':
        return 'file';
      case 'ods':
        return 'file-spreadsheet';
      default:
        return 'file';
    }
  };

  return (
    <div className="dashboard">
      {/* Welcome / Command Area */}
      <section className="dashboard__command" aria-labelledby="command-heading">
        <div className="dashboard__command-header">
          <h1 id="command-heading" className="dashboard__title">Welcome to MeMind</h1>
          <p className="dashboard__subtitle">
            Tell MeMind what you want to do with your spreadsheet
          </p>
        </div>

        <form className="dashboard__command-form" onSubmit={handleCommandSubmit}>
          <div className="dashboard__command-input-wrapper">
            <span className="dashboard__command-icon" aria-hidden="true">
              <Icon name="sparkles" size={22} />
            </span>
            <input
              type="text"
              className="dashboard__command-input input input--lg"
              placeholder='e.g., "Clean up the Q3 budget sheet and create a pivot table by department"'
              value={commandInput}
              onChange={(e) => setCommandInput(e.target.value)}
              autoComplete="off"
              aria-label="Natural language command for spreadsheet operations"
            />
          </div>
          <div className="dashboard__command-actions">
            <Button
              type="submit"
              variant="primary"
              size="lg"
              disabled={!commandInput.trim()}
              leftIcon={<Icon name="send" size={18} />}
            >
              Send Command
            </Button>
            <p className="dashboard__command-hint">
              <Icon name="info" size={14} />
              AI command processing will be available in future iterations
            </p>
          </div>
        </form>

        <div className="dashboard__command-examples" aria-label="Example commands">
          <span className="dashboard__examples-label">Try asking:</span>
          <ul className="dashboard__examples-list">
            <li>"Clean up duplicate rows in my sales data"</li>
            <li>"Create a formula to calculate YoY growth"</li>
            <li>"Make a chart showing monthly revenue trends"</li>
            <li>"Generate a summary report for Q4"</li>
          </ul>
        </div>
      </section>

      <div className="dashboard__grid">
        {/* Recent Files */}
        <section className="dashboard__section" aria-labelledby="recent-files-heading">
          <Card variant="outlined">
            <CardHeader>
              <div className="dashboard__section-header">
                <h2 id="recent-files-heading" className="dashboard__section-title">Recent Files</h2>
                <span className="dashboard__section-badge badge badge--neutral">
                  {mockFiles.length} files
                </span>
              </div>
              <p className="dashboard__section-note">
                Showing example files. Your uploaded files will appear here.
              </p>
            </CardHeader>
            <CardBody>
              <div className="dashboard__file-list" role="list">
                {mockFiles.map((file) => (
                  <article key={file.id} className="dashboard__file-item" role="listitem">
                    <div className="dashboard__file-icon" aria-hidden="true">
                      <Icon name={getFileIcon(file.type)} size={24} />
                    </div>
                    <div className="dashboard__file-info">
                      <div className="dashboard__file-name-row">
                        <h3 className="dashboard__file-name">{file.name}</h3>
                        {file.isMock && (
                          <span className="badge badge--neutral badge--mock">Example</span>
                        )}
                      </div>
                      <div className="dashboard__file-meta">
                        <span className="dashboard__file-size">{formatFileSize(file.size)}</span>
                        <span className="dashboard__file-divider" aria-hidden="true">·</span>
                        <time className="dashboard__file-date" dateTime={file.modified.toISOString()}>
                          {formatDate(file.modified)}
                        </time>
                      </div>
                    </div>
                    <Button
                      variant="ghost"
                      size="icon"
                      className="dashboard__file-action"
                      aria-label={`Open ${file.name}`}
                      onClick={() => console.log('Open file:', file.id)}
                    >
                      <Icon name="arrowRight" size={18} />
                    </Button>
                  </article>
                ))}
              </div>
              <div className="dashboard__section-footer">
                <Button variant="ghost" size="sm" leftIcon={<Icon name="folderOpen" size={14} />}>
                  Browse Files
                </Button>
                <Button variant="primary" size="sm" leftIcon={<Icon name="upload" size={14} />}>
                  Upload New
                </Button>
              </div>
            </CardBody>
          </Card>
        </section>

        {/* Quick Actions */}
        <section className="dashboard__section" aria-labelledby="quick-actions-heading">
          <Card variant="outlined">
            <CardHeader>
              <h2 id="quick-actions-heading" className="dashboard__section-title">Quick Actions</h2>
              <p className="dashboard__section-note">Common spreadsheet operations</p>
            </CardHeader>
            <CardBody>
              <div className="dashboard__actions-grid" role="list">
                {quickActions.map((action) => (
                  <article
                    key={action.id}
                    className="dashboard__action-card"
                    role="listitem"
                    onClick={() => handleQuickAction(action)}
                    tabIndex={0}
                    onKeyDown={(e) => e.key === 'Enter' && handleQuickAction(action)}
                  >
                    <div className="dashboard__action-icon" aria-hidden="true">
                      <Icon name={action.icon} size={24} />
                    </div>
                    <h3 className="dashboard__action-title">{action.label}</h3>
                    <p className="dashboard__action-description">{action.description}</p>
                  </article>
                ))}
              </div>
            </CardBody>
          </Card>
        </section>

        {/* Activity / Status */}
        <section className="dashboard__section dashboard__section--full" aria-labelledby="activity-heading">
          <Card variant="outlined">
            <CardHeader>
              <h2 id="activity-heading" className="dashboard__section-title">Activity</h2>
              <p className="dashboard__section-note">Processing status and history</p>
            </CardHeader>
            <CardBody>
              <div className="dashboard__empty-state empty-state">
                <div className="empty-state__icon" aria-hidden="true">
                  <Icon name="clock" size={48} />
                </div>
                <h3 className="empty-state__title">No activity yet</h3>
                <p className="empty-state__description">
                  Upload a file or send a command to see processing activity here.
                </p>
                <Button variant="primary" leftIcon={<Icon name="upload" size={16} />}>
                  Upload Your First File
                </Button>
              </div>
            </CardBody>
          </Card>
        </section>
      </div>
    </div>
  );
};

export default Dashboard;
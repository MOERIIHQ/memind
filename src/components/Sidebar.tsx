import { useState } from 'react';
import type { NavItem } from '../types';
import Button from './Button';
import Icon from './Icon';
import './Sidebar.css';

interface SidebarProps {
  items: NavItem[];
  activeItem: string;
  onNavigate: (id: string) => void;
  collapsed: boolean;
  onToggleCollapse: () => void;
}

const Sidebar = ({
  items,
  activeItem,
  onNavigate,
  collapsed,
  onToggleCollapse,
}: SidebarProps) => {
  const [hovered, setHovered] = useState(false);

  const isExpanded = !collapsed || hovered;

  return (
    <aside
      className={`sidebar ${collapsed ? 'sidebar--collapsed' : ''} ${hovered ? 'sidebar--hovered' : ''}`}
      onMouseEnter={() => setHovered(true)}
      onMouseLeave={() => setHovered(false)}
      role="navigation"
      aria-label="Main navigation"
    >
      <div className="sidebar__header">
        <div className="sidebar__brand">
          <Icon name="sparkles" className="sidebar__logo" size={28} />
          {isExpanded && (
            <span className="sidebar__title">MeMind</span>
          )}
        </div>
        {isExpanded && (
          <Button
            variant="ghost"
            size="icon"
            className="sidebar__toggle"
            onClick={onToggleCollapse}
            aria-label="Collapse sidebar"
            aria-expanded={!collapsed}
          >
            <Icon name={collapsed ? 'chevronRight' : 'chevronLeft'} size={18} />
          </Button>
        )}
      </div>

      <nav className="sidebar__nav" aria-label="Primary">
        <ul className="sidebar__list" role="list">
          {items.map((item) => (
            <li key={item.id} className="sidebar__item">
              <button
                className={`sidebar__link ${activeItem === item.id ? 'sidebar__link--active' : ''} ${item.disabled ? 'sidebar__link--disabled' : ''}`}
                onClick={() => !item.disabled && onNavigate(item.id)}
                disabled={item.disabled}
                role="menuitem"
                aria-current={activeItem === item.id ? 'page' : undefined}
                title={collapsed ? item.label : undefined}
              >
                <span className="sidebar__icon" aria-hidden="true">
                  <Icon name={item.icon} size={20} />
                </span>
                {isExpanded && (
                  <>
                    <span className="sidebar__label">{item.label}</span>
                    {item.badge && (
                      <span className="sidebar__badge badge badge--primary">{item.badge}</span>
                    )}
                  </>
                )}
              </button>
            </li>
          ))}
        </ul>
      </nav>

      <div className="sidebar__footer">
        {isExpanded && (
          <div className="sidebar__version">
            <span className="sidebar__version-label">Version</span>
            <span className="sidebar__version-number">0.1.0</span>
          </div>
        )}
      </div>
    </aside>
  );
};

export default Sidebar;
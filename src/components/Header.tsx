import { useState } from 'react';
import { useRef } from 'react';
import Button from './Button';
import Icon from './Icon';
import './Header.css';

interface HeaderProps {
  onMenuClick: () => void;
}

const Header = ({ onMenuClick }: HeaderProps) => {
  const [searchQuery, setSearchQuery] = useState('');
  const searchInputRef = useRef<HTMLInputElement>(null);

  const handleSearch = (e: React.FormEvent) => {
    e.preventDefault();
    // Search functionality will be implemented later
    console.log('Search:', searchQuery);
  };

  return (
    <header className="header" role="banner">
      <div className="header__left">
        <Button
          variant="ghost"
          size="icon"
          className="header__menu-btn"
          onClick={onMenuClick}
          aria-label="Toggle navigation menu"
        >
          <Icon name="menu" size={24} />
        </Button>
      </div>

      <div className="header__center">
        <form className="header__search" onSubmit={handleSearch} role="search">
          <label htmlFor="global-search" className="sr-only">
            Search MeMind
          </label>
          <span className="header__search-icon" aria-hidden="true">
            <Icon name="search" size={18} />
          </span>
          <input
            ref={searchInputRef}
            type="search"
            id="global-search"
            className="header__search-input input"
            placeholder="Search files, commands, help..."
            value={searchQuery}
            onChange={(e) => setSearchQuery(e.target.value)}
            autoComplete="off"
          />
          {searchQuery && (
            <Button
              type="button"
              variant="ghost"
              size="icon"
              className="header__search-clear"
              onClick={() => setSearchQuery('')}
              aria-label="Clear search"
            >
              <Icon name="x" size={16} />
            </Button>
          )}
        </form>
      </div>

      <div className="header__right">
        <div className="header__actions">
          <Button
            variant="ghost"
            size="icon"
            className="header__action-btn"
            aria-label="Notifications"
          >
            <Icon name="alertCircle" size={20} />
          </Button>
          <Button
            variant="ghost"
            size="icon"
            className="header__action-btn"
            aria-label="Help"
          >
            <Icon name="info" size={20} />
          </Button>
          <div className="header__divider" aria-hidden="true" />
          <Button
            variant="primary"
            size="sm"
            className="header__upload-btn"
            leftIcon={<Icon name="upload" size={16} />}
          >
            Upload
          </Button>
        </div>

        <div className="header__user">
          <Button
            variant="ghost"
            size="icon"
            className="header__user-btn"
            aria-label="User menu"
            aria-expanded="false"
            aria-haspopup="true"
          >
            <div className="header__avatar avatar" aria-hidden="true">M</div>
          </Button>
        </div>
      </div>
    </header>
  );
};

export default Header;
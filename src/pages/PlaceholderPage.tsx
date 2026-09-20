import Card, { CardHeader, CardBody } from '../components/Card';
import Icon from '../components/Icon';
import Button from '../components/Button';
import type { PlaceholderPagePrimaryAction } from '../types';
import './PlaceholderPage.css';

interface PlaceholderPageProps {
  title: string;
  description: string;
  icon: string;
  features: string[];
  primaryAction?: PlaceholderPagePrimaryAction;
}

const PlaceholderPage = ({
  title,
  description,
  icon,
  features,
  primaryAction,
}: PlaceholderPageProps) => {
  return (
    <div className="placeholder-page">
      <header className="placeholder-page__header">
        <div className="placeholder-page__icon" aria-hidden="true">
          <Icon name={icon} size={48} />
        </div>
        <h1 className="placeholder-page__title">{title}</h1>
        <p className="placeholder-page__description">{description}</p>
      </header>

      <section className="placeholder-page__features" aria-labelledby="features-heading">
        <h2 id="features-heading" className="sr-only">Planned Features</h2>
        <div className="placeholder-page__features-grid">
          {features.map((feature, index) => (
            <Card key={index} variant="outlined" className="placeholder-page__feature-card">
              <CardBody>
                <div className="placeholder-page__feature-icon" aria-hidden="true">
                  <Icon name="check" size={20} />
                </div>
                <p className="placeholder-page__feature-text">{feature}</p>
              </CardBody>
            </Card>
          ))}
        </div>
      </section>

      <div className="placeholder-page__status">
        <Card variant="outlined">
          <CardHeader>
            <h3 className="placeholder-page__status-title">Development Status</h3>
          </CardHeader>
          <CardBody>
            <div className="placeholder-page__status-content">
              <div className="placeholder-page__status-indicator">
                <span className="placeholder-page__status-dot status--planned" aria-hidden="true"></span>
                <span>Planned for future iteration</span>
              </div>
              <p className="placeholder-page__status-note">
                This module is part of the MeMind roadmap but has not been implemented yet.
                The navigation structure is in place for when development begins.
              </p>
            </div>
          </CardBody>
        </Card>
      </div>

      {primaryAction && (
        <div className="placeholder-page__action">
          <Button variant="primary" size="lg" onClick={primaryAction.primaryAction}>
            {primaryAction.label}
          </Button>
        </div>
      )}
    </div>
  );
};

export default PlaceholderPage;
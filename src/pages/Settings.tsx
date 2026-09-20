import PlaceholderPage from './PlaceholderPage';

const Settings = () => {
  return (
    <PlaceholderPage
      title="Settings"
      description="Configure MeMind to match your workflow. Manage preferences, integrations, and account settings."
      icon="settings"
      features={[
        "Appearance: light/dark/system theme",
        "Default file handling preferences",
        "AI model configuration and API keys",
        "Keyboard shortcuts customization",
        "Data privacy and local processing options",
        "Export format defaults",
      ]}
    />
  );
};

export default Settings;
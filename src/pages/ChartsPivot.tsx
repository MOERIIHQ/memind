import PlaceholderPage from './PlaceholderPage';

const ChartsPivot = () => {
  return (
    <PlaceholderPage
      title="Charts & Pivot"
      description="Create insightful visualizations and pivot tables from your spreadsheet data. MeMind will suggest the best chart types and generate pivot analyses automatically."
      icon="chart-bar"
      features={[
        "Automatic chart type recommendation",
        "Pivot table generation from natural language",
        "Interactive charts with drill-down capability",
        "Multiple chart types: bar, line, area, scatter, heatmap",
        "Conditional formatting and data bars",
        "Export charts as images or embeddable components",
      ]}
      primaryAction={{
        label: 'Create Visualization',
        primaryAction: () => console.log('Create visualization'),
      }}
    />
  );
};

export default ChartsPivot;
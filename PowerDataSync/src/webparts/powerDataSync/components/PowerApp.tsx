import * as React from 'react';
import PowerDashboard from './Dashboard/PowerDashboard';
import PowerDataSync from './PowerDataSync';
import { WebPartContext } from '@microsoft/sp-webpart-base';

// 1. UPDATE INTERFACE
export interface PowerAppProps {
  siteUrl: string;
  context: WebPartContext; // Strongly typed context
  version?: string;
  metricsListTitle: string;
  metricsLibTitle: string;
  showUserAlerts: boolean; 
  showHiddenLists: boolean;
}

export interface PowerAppState {
  route: 'dashboard' | 'wizard';
  resetKey: number;
  resumeJobId?: number;
}

export default class PowerApp extends React.Component<PowerAppProps, PowerAppState> {
  constructor(props: PowerAppProps) {
    super(props);
    this.state = { route: 'dashboard', resetKey: 0 };
  }

  private goWizard = () => this.setState({ route: 'wizard', resumeJobId: undefined });
  private goDash = () => this.setState({ route: 'dashboard', resumeJobId: undefined });

  private handleResumeJob = (id: number) => {
    this.setState((prev) => ({
      route: 'wizard',
      resetKey: prev.resetKey + 1,
      resumeJobId: id
    }));
  };

  private runAnother = () => {
    this.setState((prev) => ({
      route: 'wizard',
      resetKey: prev.resetKey + 1,
      resumeJobId: undefined
    }));
  };

  public render() {
    const { siteUrl, context, version, metricsListTitle, metricsLibTitle, showUserAlerts, showHiddenLists } = this.props;

    if (this.state.route === 'dashboard') {
      return (
        <PowerDashboard
          siteUrl={siteUrl}
          context={context} // ADDED: Passing context down for PnP v4
          metricsListTitle={metricsListTitle}
          onNewJob={this.goWizard}
          onResumeJob={this.handleResumeJob}
        />
      );
    }

    return (
      <PowerDataSync
        key={this.state.resetKey}
        siteUrl={siteUrl}
        context={context}
        version={version}
        metricsListTitle={metricsListTitle}
        metricsLibTitle={metricsLibTitle}
        onExitToDashboard={this.goDash}
        onRunAnother={this.runAnother}
        resumeJobId={this.state.resumeJobId}
        showUserAlerts={showUserAlerts} 
        showHiddenLists={showHiddenLists}
      />
    );
  }
}
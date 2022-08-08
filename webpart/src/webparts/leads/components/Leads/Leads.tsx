import * as React from 'react';
import styles from './Leads.module.scss';
import { ILeadsProps, ILeadsState, LeadView } from '.';
import { LeadCardActions, LeadCardPreview } from '..';
import { Dropdown } from 'office-ui-fabric-react/lib/Dropdown';
import { DatePicker } from 'office-ui-fabric-react/lib/DatePicker';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { css } from 'office-ui-fabric-react/lib/Utilities';
import { DocumentCard, DocumentCardTitle, DocumentCardActivity, DocumentCardLocation, IDocumentCardActivityPerson } from 'office-ui-fabric-react/lib/DocumentCard';
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { ImageFit } from 'office-ui-fabric-react/lib/Image';
// import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { HttpClient, HttpClientResponse } from '@microsoft/sp-http';
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";
import { SampleLeads } from '../../../../SampleLeads';
import { ILeadComment, ILead, IPerson } from '../../../../Lead';
import { BaseButton, Button, DefaultButton, PrimaryButton } from 'office-ui-fabric-react/lib/Button';

export class Leads extends React.Component<ILeadsProps, ILeadsState> {
  public constructor(props: ILeadsProps) {
    super(props);

    this.state = {
      loading: false,
      error: undefined,
      leads: [],
      reminderCreating: false,
      reminderDialogVisible: false,
      submitCardDialogVisible: false,
      view: typeof props.view !== 'undefined' ? props.view : LeadView.new
    };
  }

  private _loadLeads(view: LeadView): void {
    if (this.props.needsConfiguration) {
      return;
    }

    this.setState({
      loading: true,
      error: undefined
    });

    this
      ._getData()
      .then((data: ILead[]): void => {
        let leads: ILead[] = data;
        switch (view) {
          case LeadView.new:
            // eslint-disable-next-line
            leads = data.sort((a, b) => (b.createdOn as any) - (a.createdOn as any));
            break;
          case LeadView.mostProbable:
            leads = data.sort((a, b) => b.percentComplete - a.percentComplete);
            break;
          case LeadView.recentComments:
            // eslint-disable-next-line
            leads = data.sort((a, b) => (this._getLastCommentDate(b) as any) - (this._getLastCommentDate(a) as any));
            break;
          case LeadView.requireAttention:
            leads = data.filter(l => l.requiresAttention === true);
            break;
        }

        this.setState({
          loading: false,
          leads: leads
        });
      }, (error): void => {
        this.setState({
          loading: false,
          error: error
        });
      });
  }

  private _getData(): Promise<ILead[]> {
    if (this.props.demo) {
      return Promise.resolve(SampleLeads.leads);
    }
    else {
      return this.props.httpClient.get(this.props.leadsApiUrl, HttpClient.configurations.v1)
        .then((res: HttpClientResponse) => res.json());
    }
  }

  private _getLastCommentDate(lead: ILead): Date {
    let date: Date = new Date(0);

    if (lead.comments && lead.comments.length > 0) {
      date = new Date(lead.comments[lead.comments.length - 1].date);
    }

    return date;
  }

  private _viewChanged = (item): void => {
    this.setState({
      view: item.data
    });
    this._loadLeads(item.data);
  }

  private _leadClicked = (ev?: React.SyntheticEvent<HTMLElement>): void => {
    if (!this.props.teamsContext ||
      !this.props.host ||
      !this.props.host._teamsManager ||
      !this.props.host._teamsManager._appContext ||
      !this.props.host._teamsManager._appContext.applicationName) {
      return;
    }

    const host: string = this.props.host._teamsManager._appContext.applicationName;
    if (host !== 'TeamsTaskModuleApplication') {
      return;
    }

    const leadElement: HTMLElement = ev.currentTarget;
    const leadId: string = leadElement.dataset.leadid;
    if (!leadId) {
      return;
    }

    const selectedLead: ILead[] = this.state.leads.filter(lead => lead.id === leadId);
    if (selectedLead.length < 1) {
      return;
    }

    this.setState({ submitCardDialogVisible: true });
    this.props.teamsContext.teamsJs.tasks.submitTask(selectedLead[0]);
  }

  private _leadShared = (): void => {
  }

  private _leadFollowed = (): void => {
  }

  private _showCreateLeadReminderDialog = (ev?: React.MouseEvent<HTMLDivElement | HTMLAnchorElement | HTMLButtonElement | BaseButton | Button, MouseEvent>): void => {
    // eslint-disable-next-line
    const leadElement: any = ev.currentTarget;
    const leadId: string = leadElement.dataset.leadid;
    if (!leadId) {
      return;
    }

    const selectedLead: ILead[] = this.state.leads.filter(lead => lead.id === leadId);
    if (selectedLead.length < 1) {
      return;
    }

    this.setState({
      selectedLead: selectedLead[0],
      reminderDialogVisible: true,
      reminderCreating: false,
      reminderCreatingResult: undefined,
      reminderDate: new Date()
    });
  }

  private _createLeadReminder = (): void => {
    this.setState({
      reminderCreating: true,
      reminderCreatingResult: undefined
    });

    const dueDate: Date = this.state.reminderDate;
    dueDate.setHours(9);
    const reminder: string = `Review lead ${this.state.selectedLead.title}`;

    this.props.msGraphClient
      .api('me/todo/lists')
      .get()
      .then((lists: { value: { id: string; wellknownListName: string; }[] }): Promise<void> => {
        if (lists.value.length < 1) {
          return Promise.reject('No To Do lists found');
        }

        let defaultList: { id: string; wellknownListName: string; }[] = lists.value.filter(list => list.wellknownListName === 'defaultList');
        if (defaultList.length < 1) {
          defaultList = [lists.value[0]];
        }

        if (defaultList.length !== 1) {
          return Promise.reject('Default To Do list not found');
        }

        return this.props.msGraphClient
          .api(`me/todo/lists/${defaultList[0].id}/tasks`)
          .post({
            title: reminder,
            body: {
              content: `<html>\r\n<head>\r\n<meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\">\r\n<meta content=\"text/html; charset=us-ascii\">\r\n</head>\r\n<body>\r\n<p>${location.href}</p>\r\n</body>\r\n</html>\r\n`,
              'contentType': 'html'
            },
            reminderDateTime: {
              dateTime: dueDate.toISOString(),
              timeZone: 'UTC'
            }
          });
      })
      .then((): void => {
        this.setState({
          reminderCreating: false,
          reminderCreatingResult: 'OK'
        });
      }, (err): void => {
        this.setState({
          reminderCreating: false,
          reminderCreatingResult: err
        });
      });
  }

  private _closeCreateReminderDialog = (): void => {
    this.setState({
      reminderDialogVisible: false
    });
  }

  public componentDidMount(): void {
    this._loadLeads(this.state.view);
  }

  // eslint-disable-next-line
  public componentDidUpdate(prevProps: ILeadsProps, prevState: ILeadsState, snapshot?: any): void {
    if (this.props.demo !== prevProps.demo) {
      this._loadLeads(this.state.view);
    }
  }

  // eslint-disable-next-line
  private _getCommentsForCard(comments: ILeadComment[]): any {
    return comments.map(c => {
      return {
        name: c.comment,
        url: '#',
        previewImageSrc: '',
        iconSrc: SampleLeads.getPhotoUrl(c.createdBy.email),
        imageFit: ImageFit.cover,
        width: 318,
        height: 196
      };
    });
  }

  private _getDisplayDate(d: string): string {
    const date: Date = new Date(d);
    let displayDate: string = date.toLocaleDateString();
    const dateYear: number = date.getFullYear();
    const now: Date = new Date();

    if (dateYear === now.getFullYear()) {
      displayDate = displayDate.replace(dateYear.toString(), '').replace(/(^[^\d]|[^\d]$)/g, '');
    }

    return displayDate;
  }

  private _getDistinctContributors(createdBy: IPerson, comments: ILeadComment[]): IPerson[] {
    const contributors: IPerson[] = [];

    comments.forEach(c => {
      if (c.createdBy.email === createdBy.email) {
        return;
      }

      for (let i: number = 0; i < contributors.length; i++) {
        if (contributors[i].email === c.createdBy.email) {
          return;
        }
      }

      contributors.push(c.createdBy);
    });

    return contributors;
  }

  public render(): React.ReactElement<ILeadsProps> {
    const { error, loading, leads, view, submitCardDialogVisible, reminderDialogVisible, reminderCreating, reminderCreatingResult, reminderDate } = this.state;
    const { needsConfiguration } = this.props;

    if (needsConfiguration) {
      return <div className={css(styles.leads, 'ms-Fabric')}>
        <Placeholder
          iconName='Chart'
          iconText='Configure your environment'
          description='The required LeadsApiUrl tenant property is not configured. Please configure the property before using this web part.' />
      </div>;
    }

    return (
      <div className={css(styles.leads, 'ms-Fabric')}>
        {loading &&
          <Spinner label='Loading Leads...' size={SpinnerSize.large} />}
        {!loading &&
          error &&
          <div>The following error has occurred while loading Leads: {this.state.error}</div>}
        {!loading &&
          !error &&
          leads.length === 0 &&
          <div>No Leads found</div>}
        {!loading &&
          !error &&
          leads.length > 0 &&
          <div>
            <div className={styles.viewWrapper}>
              <div className={styles.title}>Leads from the Lead Management System</div>
              {typeof this.props.view === 'undefined' &&
                <Dropdown
                  placeHolder='Select view'
                  className={styles.view}
                  options={
                    [
                      { key: 'new', text: 'New Leads', data: LeadView.new },
                      { key: 'mostProbable', text: 'Most probable', data: LeadView.mostProbable },
                      { key: 'recentComments', text: 'Recently commented', data: LeadView.recentComments },
                      { key: 'requireAttention', text: 'Require attention', data: LeadView.requireAttention }
                    ]
                  }
                  selectedKey={LeadView[view]}
                  onChanged={this._viewChanged} />
              }
            </div>
            <div className={styles.cards}>
              {
                leads.map(l =>
                  <DocumentCard onClick={this._leadClicked} className={styles.card} key={l.id} data-leadid={l.id}>
                    <div className={styles.cardContents}>
                      <div>
                        <LeadCardPreview previewItems={this._getCommentsForCard(l.comments.reverse())} />
                        {l.requiresAttention === true &&
                          <div className={styles.urgent} />}
                        <DocumentCardLocation location={l.account} locationHref='#' />
                        <DocumentCardTitle title={l.title} shouldTruncate={true} />
                        <div className={styles.titleSecondary}><DocumentCardTitle title={l.description!} shouldTruncate={true} /></div>
                      </div>
                      <div>
                        <DocumentCardActivity
                          activity={`Created ${this._getDisplayDate(l.createdOn)}`}
                          people={
                            ([] as IDocumentCardActivityPerson[])
                              .concat([{ name: l.createdBy.name, profileImageSrc: SampleLeads.getPhotoUrl(l.createdBy.email) }])
                              .concat(this._getDistinctContributors(l.createdBy, l.comments).map(c => {
                                return {
                                  name: c.name,
                                  profileImageSrc: SampleLeads.getPhotoUrl(c.email)
                                };
                              }))
                          }
                        />
                        <LeadCardActions
                          actions={
                            [
                              {
                                className: 'icon share',
                                onClick: this._leadShared,
                                ariaLabel: 'Share Lead'
                              },
                              {
                                className: 'icon view',
                                onClick: this._leadFollowed,
                                ariaLabel: 'Follow Lead'
                              },
                              {
                                className: 'icon alarmClock',
                                onClick: this._showCreateLeadReminderDialog,
                                ariaLabel: 'Remind me about this lead',
                                'data-leadid': l.id
                              } as any
                            ]
                          }
                          percentComplete={l.percentComplete}
                          change={l.change} />
                      </div>
                    </div>
                  </DocumentCard>
                )
              }
            </div>
            <Dialog
              hidden={!submitCardDialogVisible}
              dialogContentProps={{
                type: DialogType.normal,
                title: 'Loading data'
              }}
              styles={{
                main: [{
                  selectors: {
                    ['@media (min-width: 480px)']: {
                      minWidth: '500px',
                      minHeight: '200px'
                    }
                  }
                }]
              }}
              modalProps={{
                isBlocking: true,
                dragOptions: undefined,
              }}>
              <Spinner label='Loading data from the Lead Management System...' labelPosition='right' size={SpinnerSize.large} style={{ margin: '2em auto' }} />
            </Dialog>
            <Dialog
              hidden={!reminderDialogVisible}
              dialogContentProps={{
                type: DialogType.normal,
                title: 'Remind me about this lead'
              }}
              styles={{
                main: [{
                  selectors: {
                    ['@media (min-width: 480px)']: {
                      minWidth: '500px',
                      minHeight: '200px'
                    }
                  }
                }]
              }}
              modalProps={{
                isBlocking: true,
                dragOptions: undefined,
              }}>
              <div>
                {!reminderCreating &&
                  <div>
                    <DatePicker
                      disabled={reminderCreatingResult === 'OK'}
                      label='Remind me about this lead on'
                      value={this.state.reminderDate}
                      onSelectDate={(date: Date | null | undefined) => this.setState({ reminderDate: date })}
                    />
                    {reminderCreatingResult === 'OK' &&
                      <div style={{ color: 'green' }}>Reminder successfully created</div>}
                    {reminderCreatingResult &&
                      reminderCreatingResult !== 'OK' &&
                      <div style={{ color: 'red' }}>An error has occurred while creating the reminder: {reminderCreatingResult}</div>}
                  </div>
                }
                {reminderCreating &&
                  <Spinner label='Creating reminder...' labelPosition='right' size={SpinnerSize.large} style={{ margin: '2em auto' }} />}
              </div>
              <DialogFooter>
                <DefaultButton onClick={() => this._closeCreateReminderDialog()} text={reminderCreatingResult === 'OK' ? 'Close' : 'Cancel'} disabled={reminderCreating} />
                <PrimaryButton onClick={() => this._createLeadReminder()} text='Create reminder' disabled={typeof reminderDate === 'undefined' || reminderCreatingResult === 'OK'} />
              </DialogFooter>
            </Dialog>
          </div>
        }
      </div>
    );
  }
}

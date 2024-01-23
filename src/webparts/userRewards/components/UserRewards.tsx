import * as React from 'react';
import styles from './UserRewards.module.scss';
import type { IUserRewardsProps } from './IUserRewardsProps';
import { Persona, PersonaSize, Spinner } from '@fluentui/react';
import { LivePersona } from '@pnp/spfx-controls-react/lib/LivePersona';

interface IUserRewardsState {
  items: any;
  isLoading: boolean;
  error: string | null;
}
export default class UserRewards extends React.Component<IUserRewardsProps, IUserRewardsState> {
  constructor(props: IUserRewardsProps | Readonly<IUserRewardsProps>) {
    super(props);

    this.state = {
      items: new Array<any>(),
      isLoading: true,
      error: null,
    }
  }

  public componentDidMount(): void {
    this.loadData().then(data => {
      this.setState({
        items: data,
        isLoading: false
      });
    }).catch((ex) => {
      this.setState({
        error: ex.message,
        isLoading: false
      });
    });
    this.getItem();
  }

  public getItem(): void {
    const siteUrl = "https://xvzms.sharepoint.com/";
    const libraryName = "Documents";
    const fileName = "TestDocument.docx";

    // Get the current item (document) ID
    const getItemUrl = `${siteUrl}/_api/web/lists/getbytitle('${libraryName}')/items?$filter=FileLeafRef eq '${fileName}'&$select=ID`;


    fetch(getItemUrl, {
      method: 'GET',
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-Type': 'application/json',
      },
    })
      .then(response => response.json())
      .then(data => {
        if (data.value.length > 0) {
          const itemId = data.value[0].ID;

          // Get the version history of the item
          const getVersionHistoryUrl = `${siteUrl}/_api/web/lists/getbytitle('${libraryName}')/items(${itemId})/versions`;

          fetch(getVersionHistoryUrl, {
            method: 'GET',
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'Content-Type': 'application/json',
            },
          })
            .then(response => response.json())
            .then(versionData => {
              if (versionData.value.length > 0) {
                const versionId = versionData.value[0].VersionId;

                // Download the specific version of the document
                const downloadUrl = `${siteUrl}/_layouts/15/download.aspx?SourceUrl=${encodeURIComponent(`/${libraryName}/${fileName}`)}&FldEdit=0&ver=${versionId}`;
                console.log("Download URL: " + downloadUrl);
                window.open(downloadUrl, '_blank');
              } else {
                console.error('No version history found for the document');
              }
            })
            .catch(error => {
              console.error('Error fetching version history:', error);
            });
        } else {
          console.error('Document not found');
        }
      })
      .catch(error => {
        console.error('Error fetching item ID:', error);
      });
  }

  public render(): React.ReactElement<IUserRewardsProps> {
    const { context } = this.props;
    const { items, isLoading, error } = this.state;

    return (
      <section className={styles.userRewards}>
        <div>
          <h1>User Rewards</h1>
          {error && <div>{error}</div>}
          {
            isLoading ?
              <div><Spinner /></div>
              :
              <div>
                {items.map((i: any) => {
                  // eslint-disable-next-line react/jsx-key
                  return <div>
                    <LivePersona
                      upn={i.employee.Email}
                      template={
                        <>
                          <Persona
                            imageUrl={`https://scisoft.sharepoint.com/sites/Demo/LMS/_layouts/15/userphoto.aspx?size=M&username=${i.employee.Email}`}
                            // imageUrl={`https://xvzms.sharepoint.com/_layouts/15/userphoto.aspx?size=M&username=${i.employee.Email}`}
                            imageAlt={i.employee.Email}
                            text={i.employee.Title}
                            size={PersonaSize.regular}
                          />
                        </>
                      }
                      serviceScope={context.serviceScope}
                      disableHover={false}
                    />
                    <div>
                      <div><span>Reward: </span><span>{i.reward}</span></div>
                    </div>
                  </div>;
                })
                }
              </div>
          }
        </div>
      </section>
    );
  }

  private loadData(): Promise<Array<any>> {
    return this.props.service.retrieveListItems().then((result: Array<any>) => {
      return result;
    });
  }
}

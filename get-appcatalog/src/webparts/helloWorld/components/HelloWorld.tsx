import * as React from 'react';
import styles from './HelloWorld.module.scss';
import { IHelloWorldProps } from './IHelloWorldProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from '@pnp/sp';
import "@pnp/sp/appcatalog";
import "@pnp/sp/webs";

export default class HelloWorld extends React.Component<IHelloWorldProps, {}> {
  public state = { appCatWeb: null }

  public async getAppcatalogURL(): Promise<string> {
    const appCatWeb = await sp.getTenantAppCatalogWeb();
    const { Url } = await appCatWeb()
    return Url;
  }

  public async componentDidMount() {
    const url =  await this.getAppcatalogURL();
    this.setState({...this.state, appCatWeb:url})
  }

  public render(): React.ReactElement<IHelloWorldProps> {
    return (
      <div className={styles.helloWorld}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <span className={styles.title}>{this.state.appCatWeb}</span>
              <p className={styles.subTitle}>Customize SharePoint experiences using Web Parts.</p>
              <p className={styles.description}>{escape(this.props.description)}</p>
              <a href="https://aka.ms/spfx" className={styles.button}>
                <span className={styles.label}>Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>
    );
  }
}

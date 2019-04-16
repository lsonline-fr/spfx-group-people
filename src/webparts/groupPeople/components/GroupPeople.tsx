import * as React from 'react';
import styles from './GroupPeople.module.scss';
import { IGroupPeopleProps } from './IGroupPeopleProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class GroupPeople extends React.Component<IGroupPeopleProps, {}> {
  public render(): React.ReactElement<IGroupPeopleProps> {
    return (
      <div className={ styles.groupPeople }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>{escape(this.props.title)}</span>
              <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>
              <p className={ styles.description }>description</p>
              <a href="https://aka.ms/spfx" className={ styles.button }>
                <span className={ styles.label }>Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>
    );
  }
}

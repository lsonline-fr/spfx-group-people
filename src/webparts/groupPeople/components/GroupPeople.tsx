import * as React from 'react';
import styles from './GroupPeople.module.scss';
import { IGroupPeopleProps } from './IGroupPeopleProps';

import { Persona, PersonaSize } from 'office-ui-fabric-react/lib/Persona';

export default class GroupPeople extends React.Component<IGroupPeopleProps, {}> {

  private _toggleTitle: string = '';

  constructor(props: IGroupPeopleProps) {
    super(props);
    this._toggleTitle = props.displayTitle ? '' : styles.hidden;
  }

  public render(): JSX.Element {
    this._toggleTitle = this.props.displayTitle ? '' : styles.hidden;
    return (
      <div className={styles.groupPeople}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <div className={[styles.title, this._toggleTitle].join(' ')} role="heading">{this.props.title}</div>
                {this.props.users.map(function (u:any) {
                  return (<div className={styles.personaTile}><Persona
                    text={u.DisplayName} 
                    secondaryText={u.Title}
                    imageUrl={u.PictureUrl} 
                    size={PersonaSize.size48}
                    className={styles.persona}
                  /></div>)
                })}
            </div>
          </div>
        </div>
      </div>
    );
  }
}

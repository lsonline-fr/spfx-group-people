import * as React from 'react';
import styles from './GroupPeople.module.scss';
import { IGroupPeopleProps } from './IGroupPeopleProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { IPersonaSharedProps, Persona, PersonaSize } from 'office-ui-fabric-react/lib/Persona';

export default class GroupPeople extends React.Component<IGroupPeopleProps, {}> {

  constructor(props: IGroupPeopleProps) {
    super(props);
  }

  public render(): JSX.Element {
    return (
      <div className={styles.groupPeople}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <div className={styles.title} role="heading">{escape(this.props.title)}</div>
              <div>
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
      </div>
    );
  }
}

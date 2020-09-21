import * as React from 'react';
import styles from './GroupPeople.module.scss';
import { IGroupPeopleProps } from './IGroupPeopleProps';

import { Persona } from 'office-ui-fabric-react/lib/Persona';
import PeopleCard from '../models/PeopleCard';

import * as strings from 'GroupPeopleWebPartStrings';

/** Group People UI
 * @class
 * @extends
 */
export default class GroupPeople extends React.Component<IGroupPeopleProps, {}> {

  /** Toggle Title state
   * @private
   */
  private _toggleTitle: string = '';

  /** Display a message if no People to display and if don't hide webpart
   * @private
   */
  private _displayDefaultMessage: string = '';

  /** Default constructor
   * @param props 
   */
  constructor(props: IGroupPeopleProps) {
    super(props);
    this._toggleTitle = props.displayTitle ? '' : styles.hidden;
  }

  /** Default render
   * @returns HTML Template
   * @public
   */
  public render(): JSX.Element {
    this._toggleTitle = this.props.displayTitle ? '' : styles.hidden;
    this._displayDefaultMessage = (this.props.users.length == 0 && !this.props.hide) ? '' : styles.hidden;
    return (
      <div className={styles.groupPeople}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <h2 className={[styles.title, this._toggleTitle].join(' ')} role="heading">{this.props.title}</h2>
              <div className={['grpPeopleNoItem', this._displayDefaultMessage].join(' ')}>{strings.NoItemFound}</div>
              {this.props.users.map((p: PeopleCard) => {
                return (<div className={styles.personaTile} key={p.key}><Persona
                  text={p.lineOne}
                  secondaryText={p.lineTwo}
                  tertiaryText={p.lineThree}
                  imageUrl={p.image}
                  size={this.props.size}
                  className={styles.persona}
                /></div>);
              })}
            </div>
          </div>
        </div>
      </div>
    );
  }
}

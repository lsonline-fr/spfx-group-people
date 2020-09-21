import * as React from 'react';
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";

import strings from 'GroupPeopleWebPartStrings';

/** Group People PlaceHolder UI
 * @class
 * @extends
 * @exports
 */
export default class GroupPeoplePlaceHolder extends React.Component<any> {
    
    /** Default Render
     * @returns HTML Template
     * @public
     */
    public render(): JSX.Element {
        return (
            <Placeholder
                iconName='Group'
                iconText={strings.PlaceHolderHeader}
                description={strings.PlaceHolderText} />
        );
    }
}

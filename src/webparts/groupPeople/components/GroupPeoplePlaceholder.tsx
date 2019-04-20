import * as React from 'react';
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";

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
                iconText='Group People'
                description='Display the members of a target SharePoint group' />
        );
    }
}

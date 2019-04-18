import * as React from 'react';
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";

export default class GroupPeoplePlaceHolder extends React.Component<any> {
    
    public render(): JSX.Element {
        return (
            <Placeholder
                iconName='Group'
                iconText='Group People'
                description='Display the members of a target SharePoint group' />
        );
    }
}

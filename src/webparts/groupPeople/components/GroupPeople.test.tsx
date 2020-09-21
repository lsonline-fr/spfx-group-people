/// <reference types="jest" />

import * as React from 'react';
import { expect } from 'chai';
import { configure, mount, ReactWrapper } from 'enzyme';
import Adapter from 'enzyme-adapter-react-16';

configure({ adapter: new Adapter() });

import GroupPeople from './GroupPeople';
import PeopleCard from '../models/PeopleCard';
import { IGroupPeopleProps } from './IGroupPeopleProps';

describe('Group People Render', () => {
    let reactComponent: ReactWrapper<IGroupPeopleProps, {}>;

    afterEach(() => {
        reactComponent.unmount();
    });

    it('should root web part element exists', () => {
        reactComponent = mount(React.createElement(
            GroupPeople,
            {
                title: 'SharePoint Group Title',
                users: new Array,
                size: 13,
                displayTitle: true,
                hide: false
            }
          ));
        let cssSelector: string = '.groupPeople';

        const element = reactComponent.find(cssSelector);
        expect(element.length).to.be.greaterThan(0);
    });

    it('should render the title of the SharePoint group', () => {
        reactComponent = mount(React.createElement(
            GroupPeople,
            {
                title: 'SharePoint Group Title',
                users: new Array,
                size: 13,
                displayTitle: false,
                hide: false
            }
          ));
        let cssSelector: string = 'h2';

        const element = reactComponent.find(cssSelector);
        expect(element.text()).to.be.equals('SharePoint Group Title');
    });

    it('should hide the title of the webpart', () => {
        reactComponent = mount(React.createElement(
            GroupPeople,
            {
                title: 'SharePoint Group Title',
                users: new Array,
                size: 13,
                displayTitle: false,
                hide: false
            }
          ));
        let cssSelector: string = 'h2.hidden';

        const element = reactComponent.find(cssSelector);
        expect(element.length).to.be.greaterThan(0);
    });

    it('should hide the default message when no user', () => {
        reactComponent = mount(React.createElement(
            GroupPeople,
            {
                title: 'SharePoint Group Title',
                users: new Array,
                size: 13,
                displayTitle: true,
                hide: false
            }
          ));
        let cssSelector: string = 'div.grpPeopleNoItem.hidden';

        const element = reactComponent.find(cssSelector);
        expect(element.length).to.be.equals(0);
    });

    it('should render the default message when no user', () => {
        reactComponent = mount(React.createElement(
            GroupPeople,
            {
                title: 'SharePoint Group Title',
                users: new Array,
                size: 13,
                displayTitle: true,
                hide: true
            }
          ));
        let cssSelector: string = 'div.grpPeopleNoItem.hidden';

        const element = reactComponent.find(cssSelector);
        expect(element.length).to.be.equals(1);
    });

    describe('Group People Render with users', () => {
        beforeEach(() => {
            reactComponent = mount(React.createElement(
                GroupPeople,
                {
                    title: 'SharePoint Group Title',
                    users: [
                        new PeopleCard('i:0#.f|membership|user1@contoso.onmicrosoft.com', 'https://static2.sharepointonline.com/files/fabric/office-ui-fabric-react-assets/persona-female.png', 'Annie Lindqvist', 'Microsoft 365 Architect', 'Information Technology'),
                        new PeopleCard('i:0#.f|membership|user2@contoso.onmicrosoft.com', 'https://static2.sharepointonline.com/files/fabric/office-ui-fabric-react-assets/persona-male.png', 'Ted Randall', 'Microsoft 365 Developer', 'Software Development'),
                        new PeopleCard('i:0#.f|membership|user3@contoso.onmicrosoft.com', '', 'Maor Sharett', 'Microsoft 365 Developer', 'Software Development'),
                    ],
                    size: 13,
                    displayTitle: true,
                    hide: false
                }
              ));
        });

        it('should render three people cards', () => {
            let cssSelector: string = 'div.personaTile';

            const element = reactComponent.find(cssSelector);
            expect(element.length).to.be.equals(3);
        });
    });

});
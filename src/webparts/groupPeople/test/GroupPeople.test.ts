/// <reference types="jest" />

import * as React from 'react';
import { configure, mount, ReactWrapper } from 'enzyme';
import * as Adapter from 'enzyme-adapter-react-16';

configure({ adapter: new Adapter() });

import GroupPeople from '../components/GroupPeople';
import { IGroupPeopleProps } from '../components/IGroupPeopleProps';

describe('Groupe People Render', () => {

    let reactComponent: ReactWrapper<IGroupPeopleProps>;
  
    beforeEach(() => {
  
      reactComponent = mount(React.createElement(
        GroupPeople,
        {
          title: 'My Site Owners',
          users: [],
          displayTitle: true
        }
      ));
    });
  
    afterEach(() => {
      reactComponent.unmount();
    });
  
    it('should has the correct title', () => {
  
      // Arrange
      // define contains/like css selector
      let cssSelector: string = 'h2';
  
      // Act
      // find the elemet using css selector
      const text = reactComponent.find(cssSelector).text();
  
      // Assert
      expect(text).toBe('My Site Owners');  
    });

  });
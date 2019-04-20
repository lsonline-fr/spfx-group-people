/// <reference types="mocha" />

import * as React from 'react';
import { assert, expect } from 'chai';
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

    it('should do something', () => {
      assert.ok(true);
    });

    it('should render something', () => {
      expect(reactComponent.find('div')).to.not.be.null;
    });
  
    it('should has the correct title', () => {
      expect(reactComponent.find('h2').text()).to.be.equals('My Site Owners');  
    });
  });
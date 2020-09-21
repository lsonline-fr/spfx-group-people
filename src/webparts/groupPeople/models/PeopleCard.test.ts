/// <reference types="jest" />

import PeopleCard from './PeopleCard';

describe('PeopleCard', () => {
    describe("The constructor", () => {
        it('should create a PeopleCard instance with the minimum parameters', () => {
            let myPeopleCard = new PeopleCard('i:0#.f|membership|user1@contoso.onmicrosoft.com');
            expect(myPeopleCard).toEqual({
                _loginName: 'i:0#.f|membership|user1@contoso.onmicrosoft.com',
                _image: null,
                _lineOne: null,
                _lineTwo: null,
                _lineThree: null,
            });
        });

        it('should create a PeopleCard instance with all parameters', () => {
            let myPeopleCard = new PeopleCard(
                'i:0#.f|membership|user1@contoso.onmicrosoft.com',
                'https://static2.sharepointonline.com/files/fabric/office-ui-fabric-react-assets/persona-female.png',
                'Annie Lindqvist',
                'Microsoft 365 Architect',
                'Information Technology'
            );
            expect(myPeopleCard).toEqual({
                _loginName: 'i:0#.f|membership|user1@contoso.onmicrosoft.com',
                _image: 'https://static2.sharepointonline.com/files/fabric/office-ui-fabric-react-assets/persona-female.png',
                _lineOne: 'Annie Lindqvist',
                _lineTwo: 'Microsoft 365 Architect',
                _lineThree: 'Information Technology',
            });
        });

        it('should throw an error when creating a PeopleCard instance with an empty loginName', () => {
            expect(() => {
                new PeopleCard(
                    ' ',
                    'https://static2.sharepointonline.com/files/fabric/office-ui-fabric-react-assets/persona-female.png',
                    'Annie Lindqvist',
                    'Microsoft 365 Architect',
                    'Information Technology'
                );
            }).toThrow(TypeError);
        });

        it('should throw an error when creating a PeopleCard instance with a null loginName', () => {
            let loginName = null;
            expect(() => {
                new PeopleCard(
                    loginName,
                    'https://static2.sharepointonline.com/files/fabric/office-ui-fabric-react-assets/persona-female.png',
                    'Annie Lindqvist',
                    'Microsoft 365 Architect',
                    'Information Technology'
                );
            }).toThrow(TypeError);
        });
    });

    describe('Getters/Setters', () => {
        let myPeopleCard: PeopleCard;

        beforeEach(() => {
            myPeopleCard = new PeopleCard('i:0#.f|membership|user1@contoso.onmicrosoft.com');
        });

        it('should return the \'key\' when the getter \'key\' called', () => {
            expect(myPeopleCard.key).toBe(-449167891);
        });

        it('should return the \'loginName\' when the getter \'loginName\' called', () => {
            expect(myPeopleCard.loginName).toBe('i:0#.f|membership|user1@contoso.onmicrosoft.com');
        });

        it('should return the \'image\' when the getter \'image\' called', () => {
            expect(myPeopleCard.image).toBeNull;
        });

        it('should set myPeopleCard._image to the passed argument \'https://static2.sharepointonline.com/files/fabric/office-ui-fabric-react-assets/persona-female.png\'', () => {
            myPeopleCard.image = 'https://static2.sharepointonline.com/files/fabric/office-ui-fabric-react-assets/persona-female.png';
            expect(myPeopleCard.image).toBe('https://static2.sharepointonline.com/files/fabric/office-ui-fabric-react-assets/persona-female.png');
        });

        it('should return the \'lineOne\' text when the getter \'lineOne\' called', () => {
            expect(myPeopleCard.lineOne).toBeNull;
        });

        it('should set myPeopleCard._lineOne to the passed argument \'Annie Lindqvist\'', () => {
            myPeopleCard.lineOne = 'Annie Lindqvist';
            expect(myPeopleCard.lineOne).toBe('Annie Lindqvist');
        });

        it('should return the \'lineTwo\' text when the getter \'lineTwo\' called', () => {
            expect(myPeopleCard.lineTwo).toBeNull;
        });

        it('should set myPeopleCard._lineTwo to the passed argument \'Microsoft 365 Architect\'', () => {
            myPeopleCard.lineTwo = 'Microsoft 365 Architect';
            expect(myPeopleCard.lineTwo).toBe('Microsoft 365 Architect');
        });

        it('should return the \'lineThree\' text when the getter \'lineThree\' called', () => {
            expect(myPeopleCard.lineThree).toBeNull;
        });

        it('should set myPeopleCard._lineThree to the passed argument \'Information Technology\'', () => {
            myPeopleCard.lineThree = 'Information Technology';
            expect(myPeopleCard.lineThree).toBe('Information Technology');
        });
    });

});
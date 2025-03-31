const { expect } = require('chai');

describe('Clasp Test', () => {
    it('should return hello world', () => {
        const result = 'hello world';
        expect(result).to.equal('hello world');
    });
});
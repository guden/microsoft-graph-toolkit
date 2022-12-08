import { makeDecorator } from '@storybook/addons';
import { html } from 'lit-html';

export const withMockProvider = story => html`<mgt-mock-provider></mgt-mock-provider>${story()}`;

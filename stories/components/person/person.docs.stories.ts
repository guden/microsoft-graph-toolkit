/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit-html';
import { ifDefined } from 'lit-html/directives/if-defined';
import { withMockProvider } from '../../../.storybook/addons/mockProviderDecorator';

export default {
  title: 'Components / mgt-person / Presence',
  component: 'mgt-person',
  decorators: [withMockProvider]
};

export const PersonWithArgs = args => html`
  <mgt-person></mgt-person>`;

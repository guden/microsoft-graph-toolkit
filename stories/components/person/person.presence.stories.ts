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
  decorators: [withMockProvider],
  argTypes: {
    activity: { control: 'select', options: ['Available', 'OutOfOffice', 'Busy', 'DoNotDisturb', 'Away', 'Offline'] },
    availability: { control: 'select', options: ['Available', 'Busy', 'DoNotDisturb', 'Away', 'Offline'] }
  }
};

const Template = ({ personQuery, activity, availability }) => html`
  <mgt-person
    person-query=${ifDefined(personQuery)}
    show-presence
    person-presence=${JSON.stringify({ availability: ifDefined(availability), activity: ifDefined(activity) })}
  ></mgt-person>`;

export const Presence = Template.bind({});
Presence.args = {
  personQuery: 'me',
  activity: 'Available',
  availability: 'Available'
};

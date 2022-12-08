/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */
import { html } from 'lit-html';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';
import { withMockProvider } from '../../../.storybook/addons/mockProviderDecorator';
import '../../../assets/mgt.storybook.js';

export default {
  title: 'Components / mgt-login',
  component: 'mgt-login',
  decorators: [withMockProvider]
};

export const Login = () => html`<mgt-login></mgt-login>`;
Login.story = {
  parameters: { controls: { disabled: true } }
};

export const RTL = () => html`
  <body dir="rtl">
    <mgt-login></mgt-login>
  </body>
`;

// export const Events = () => html`
// <mgt-login></mgt-login>
// <script>
//   const login = document.querySelector('mgt-login');
//   login.addEventListener('loginInitiated', (e) => {
//     console.log("Login Initiated");
//   })
//   login.addEventListener('loginCompleted', (e) => {
//     console.log("Login Completed");
//   })
//   login.addEventListener('logoutInitiated', (e) => {
//     console.log("Logout Initiated");
//   })
//   login.addEventListener('logoutCompleted', (e) => {
//     console.log("Logout Completed");
//   })
// </script>
// `;

// export const localization = () => html`
//   <mgt-login></mgt-login>
//   <script>
//   import { LocalizationHelper } from '@microsoft/mgt';
//   LocalizationHelper.strings = {
//     _components: {
//       login: {
//         signInLinkSubtitle: 'Sign In 🤗',
//         signOutLinkSubtitle: 'Sign Out 🙋‍♀️'
//       },
//     }
//   }
//   </script>
// `;

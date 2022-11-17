# Utilities for SharePoint Framework Web Parts using Microsoft Graph Toolkit

[![npm](https://img.shields.io/npm/v/@microsoft/mgt-spfx-utils?style=for-the-badge)](https://www.npmjs.com/package/@microsoft/mgt-spfx-utils)

![SPFx 1.15.2](https://img.shields.io/badge/SPFx-1.15.2-green.svg?style=for-the-badge)

Helper functions to simplify lazy loading of Microsoft Graph Toolkit components when using disambiguated web components in SharePoint Framework web parts.

## Installation

To load lazy Microsoft Graph Toolkit components from the library, add the `@microsoft/mgt-spfx-utils` package and if using React, the `@microsoft/mgt-react` package as dependencies to your SharePoint Framework project:

```bash
npm install @microsoft/mgt-spfx-utils
```

or

```bash
yarn add @microsoft/mgt-spfx-utils
```

or when using React:

```bash
npm install @microsoft/mgt-spfx-utils @microsoft/mgt-react
```

or

```bash
yarn add @microsoft/mgt-spfx-utils @microsoft/mgt-react
```

> **Important:** Since a given web component tag can only be registered once these approaches **must** be used along with the `customElementHelper.withDisambiguation('foo')` approach as this allows developers to create disambiguated tag names.

By disambiguating tag names of Microsoft Graph Toolkit components, developers can use their own version of MGT rather than using the centrally deployed `@microsoft/mgt-spfx` package. This allows them to avoid colliding with other SharePoint Framework components, built by other developers. When disambiguating tag names, MGT is included in the SPFx bundle, increasing its size.

## Usage

### When using no framework webparts

When building SharePoint Framework web parts without a JavaScript framework the `@microsoft/mgt-components` library must be asynchronously loaded after configuring the disambiguation setting. The `importMgtComponentsLibrary` helper function wraps this functionality. Once the `@microsoft/mgt-components` library is loaded you can load components directly in your web part.
Below is a minimal example webpart that demonstrates how to use MGT with disambiguation in SharePoint Framework Webparts. A more complete example is available in the [No Framework Webpart Sample](../../samples/sp-mgt/src/webparts/helloWorld/HelloWorldWebPart.ts).

```ts
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { Providers } from '@microsoft/mgt-element';
import { SharePointProvider } from '@microsoft/mgt-sharepoint-provider';
import { customElementHelper } from '@microsoft/mgt-element/dist/es6/components/customElementHelper';
import { importMgtComponentsLibrary } from '@microsoft/mgt-spfx-utils';

export default class MgtWebPart extends BaseClientSideWebPart<Record<string, unknown>> {
  private _hasImportedMgtScripts = false;
  private _errorMessage = '';

  protected onInit(): Promise<void> {
    if (!Providers.globalProvider) {
      Providers.globalProvider = new SharePointProvider(this.context);
    }
    customElementHelper.withDisambiguation('foo');
    return super.onInit();
  }

  private onScriptsLoadedSuccessfully() {
    this.render();
  }

  public render(): void {
    importMgtComponentsLibrary(this._hasImportedMgtScripts, this.onScriptsLoadedSuccessfully, this.setErrorMessage);

    this.domElement.innerHTML = `
    <section class="${styles.helloWorld} ${this.context.sdks.microsoftTeams ? styles.teams : ''}">
      ${this._renderMgtComponents()}
      ${this._renderErrorMessage()}
    </section>`;
  }

  private _renderMgtComponents(): string {
    return this._hasImportedMgtScripts
      ? '<mgt-foo-login></mgt-foo-login>'
      : '';
  }

  private setErrorMessage(e?: Error): void {
    if (e) this.renderError(e);

    this._errorMessage = 'An error ocurred loading MGT scripts';
    this.render();
  }

  private _renderErrorMessage(): string {
    return this._errorMessage
      ? `<span>${this._errorMessage}</span>`
      : '';
  }
}
```

### When using React to build webparts

When building SharePoint Framework web parts using React any component that imports from the `@microsoft/mgt-react` library must be asynchronously loaded after configuring the disambiguation setting. The `lazyLoadComponent` helper function exists to facilitate using `React.lazy` and `React.Suspense` to lazy load these components from the top level webpart.
Below is a minimal example webpart that demonstrates how to use MGT with disambiguation in React based SharePoint Framework Webparts. A complete example is available in the [React SharePoint Webpart Sample](../../samples/sp-webpart/src/webparts/mgtDemo/MgtDemoWebPart.ts).

```ts
// [...] trimmed for brevity
import { Providers } from '@microsoft/mgt-element/dist/es6/providers/Providers';
import { customElementHelper } from '@microsoft/mgt-element/dist/es6/components/customElementHelper';
import { SharePointProvider } from '@microsoft/mgt-sharepoint-provider/dist/es6/SharePointProvider';
import { lazyLoadComponent } from '@microsoft/mgt-spfx-utils';

// Async import of component that imports the React Components
const MgtDemo = React.lazy(() => import('./components/MgtDemo'));

export interface IMgtDemoWebPartProps {
  description: string;
}
// set the disambiguation before initializing any webpart
customElementHelper.withDisambiguation('bar');

export default class MgtDemoWebPart extends BaseClientSideWebPart<IMgtDemoWebPartProps> {
  // set the global provider
  protected async onInit() {
    if (!Providers.globalProvider) {
      Providers.globalProvider = new SharePointProvider(this.context);
    }
  }

  public render(): void {
    const element = lazyLoadComponent(MgtDemo, { description: this.properties.description });

    ReactDom.render(element, this.domElement);
  }

  // [...] trimmed for brevity
}
```

The underlying components can then use MGT components from the `@microsoft/mgt-react` package:

```tsx
import { Person } from '@microsoft/mgt-react';

// [...] trimmed for brevity

export default class MgtReact extends React.Component<IMgtReactProps, {}> {
  public render(): React.ReactElement<IMgtReactProps> {
    return (
      <div className={ styles.mgtReact }>
        <Person personQuery="me" />
      </div>
    );
  }
}
```
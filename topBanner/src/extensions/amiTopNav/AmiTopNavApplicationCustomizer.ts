import { override } from '@microsoft/decorators';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { spfi, SPFx, SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import styles from './AmiTopNavApplicationCustomizer.module.scss';

interface ITopNavItem {
  title: string;
  url: string;
}

export interface IAmiTopNavApplicationCustomizerProperties {
  logoUrl: string;
}

export default class AmiTopNavApplicationCustomizer
  extends BaseApplicationCustomizer<IAmiTopNavApplicationCustomizerProperties> {

  private _topPlaceholder: PlaceholderContent | undefined;
  private _isCommandBarVisible: boolean = false;
  private _commandBarObserver: MutationObserver | undefined;
  private _isCommandBarShortcutBound: boolean = false;
  private readonly _topBannerSettingsListAbsoluteUrl: string = '/sites/portal/Lists/OrgSiteLinkButton';
  private _sp: SPFI | undefined;

  @override
  public onInit(): Promise<void> {
    this._sp = spfi().using(SPFx(this.context));

    this.context.placeholderProvider.changedEvent.add(this, this._renderTopNav);
    void this._renderTopNav();

    return Promise.resolve();
  }

  private async _renderTopNav(): Promise<void> {
    if (this._topPlaceholder) {
      return;
    }

    this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(
      PlaceholderName.Top,
      { onDispose: this._onDispose }
    );

    if (!this._topPlaceholder) {
      console.warn('Top placeholder was not found.');
      return;
    }

    const logoUrl =
      this.properties.logoUrl ||
      'https://ezermizionil.sharepoint.com/sites/portal/SiteAssets/image.png';

    const organizationSiteUrl = await this._loadOrganizationSiteUrlFromSettingsList();

    const navItems = await this._waitForExistingNavigationItems();
    const visibleNavItems = navItems.slice(0, 5);
    const hiddenNavItems = navItems.slice(5);

    console.log('🐰AMI NAV ITEMS:', navItems);

    const navHtml = visibleNavItems
    .map((item: ITopNavItem) => {
      return `<a href="${this._escapeHtml(item.url)}">${this._escapeHtml(item.title)}</a>`;
    })
    .join('');

    const hiddenNavHtml = hiddenNavItems
      .map((item: ITopNavItem) => {
        return `<a href="${this._escapeHtml(item.url)}">${this._escapeHtml(item.title)}</a>`;
      })
      .join('');

    const mobileNavHtml = navItems
      .map((item: ITopNavItem) => {
        return `<a href="${this._escapeHtml(item.url)}">${this._escapeHtml(item.title)}</a>`;
      })
      .join('');
        

    this._topPlaceholder.domElement.innerHTML = `
      <header class="${styles.amiTopNav}" data-ami-top-banner="true" dir="rtl">
        <div class="${styles.logoWrap}">
          <img class="${styles.logo}" src="${this._escapeHtml(logoUrl)}" alt="לוגו האתר" />
        </div>

        <nav class="${styles.navLinks}" aria-label="ניווט ראשי">
          ${navHtml}

          ${
            hiddenNavItems.length > 0
              ? `
                <div class="${styles.moreMenuWrap}">
                  <button class="${styles.moreButton}" type="button" data-ami-more-button="true">
                    ...
                  </button>

                  <div class="${styles.moreMenu}" data-ami-more-menu="true">
                    ${hiddenNavHtml}
                  </div>
                </div>
              `
              : ''
          }
        </nav>

        <a class="${styles.mainButton}" href="${this._escapeHtml(organizationSiteUrl)}">
          למעבר לאתר הארגון
        </a>
        <button class="${styles.hamburger}" type="button" data-ami-hamburger-button="true" aria-label="פתיחת תפריט">
          ☰
        </button>

        <div class="${styles.mobileMenu}" data-ami-mobile-menu="true">
          ${mobileNavHtml}
        </div>
      </header>
    `;


    const mainContent = document.querySelector(
      'section.mainContent'
    ) as HTMLElement | null;

    if (mainContent) {
      mainContent.style.marginTop = '-26px';
    }
    this._hideOriginalSharePointNavigation();
    this._bindMoreMenuEvents();
    this._setupCommandBarShortcut();
  }

private async _loadOrganizationSiteUrlFromSettingsList(): Promise<string> {
  const fallbackUrl = this.context.pageContext.web.absoluteUrl;
 if (!this._sp) {
    return fallbackUrl;
  }
  try {
    const items = await this._sp.web.getList(this._topBannerSettingsListAbsoluteUrl)
      .items.select('OrgSiteUrl')
      .top(1)();

    const orgSiteUrl = items[0]?.OrgSiteUrl?.Url?.trim();
    return orgSiteUrl || fallbackUrl;

  } catch (error) {
    console.warn('שגיאה בטעינת הקישור מהרשימה', error);
    return fallbackUrl;
  }
}
 private async _waitForExistingNavigationItems(): Promise<ITopNavItem[]> {
  const menuStateItems = await this._loadNavigationFromMenuStateApi();

  if (menuStateItems.length > 0) {
    return menuStateItems;
  }

  const topNavItems = await this._loadTopNavigationFromSharePointApi();

  if (topNavItems.length > 0) {
    return topNavItems;
  }

  for (let i = 0; i < 40; i++) {
    const items = this._getExistingNavigationItemsFromPage();

    if (items.length > 0) {
      console.log("items ", items);
      return items;
    }

    await this._delay(300);
  }

  console.warn('No SharePoint navigation links were found.');
  return [];
}

private async _loadNavigationFromMenuStateApi(): Promise<ITopNavItem[]> {
  const webUrl = this.context.pageContext.web.absoluteUrl;

  const providers = [
    'GlobalNavigationSwitchableProvider',
    'CurrentNavigationSwitchableProvider'
  ];

  for (const provider of providers) {
    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        `${webUrl}/_api/navigation/menustate?mapProviderName='${provider}'`,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        continue;
      }

      const data = await response.json();
      const nodes = data?.MenuState?.Nodes || data?.Nodes || [];

      const items: ITopNavItem[] = nodes
        .filter((node: any) => node.Title && node.SimpleUrl)
        .map((node: any) => ({
          title: node.Title,
          url: this._normalizeNavigationUrl(node.SimpleUrl)
        }));
      if (items.length > 0) {
        return items;
      }
    } catch (error) {
      console.warn(`Could not load navigation from ${provider}`, error);
    }
  }

  return [];
}

  private async _loadTopNavigationFromSharePointApi(): Promise<ITopNavItem[]> {
  const webUrl = this.context.pageContext.web.absoluteUrl;

  try {
    const response: SPHttpClientResponse = await this.context.spHttpClient.get(
      `${webUrl}/_api/web/navigation/TopNavigationBar?$select=Title,Url,Id&$orderby=Id`,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      console.warn('Could not load SharePoint top navigation from API.');
      return [];
    }

    const data = await response.json();

    return (data.value || [])
      .filter((item: { Title?: string; Url?: string }) => item.Title && item.Url)
      .map((item: { Title: string; Url: string }) => ({
        title: item.Title,
        url: this._normalizeNavigationUrl(item.Url)
      }));
  } catch (error) {
    console.warn('Error while loading SharePoint top navigation from API.', error);
    return [];
  }
}

  private _getExistingNavigationItemsFromPage(): ITopNavItem[] {
    const selectors = [
      '[data-automationid="HorizontalNav"] a[href]',
      '[data-automation-id="HorizontalNav"] a[href]',
      '[data-automationid="SiteNavigation"] a[href]',
      '[data-automation-id="SiteNavigation"] a[href]',
      '#spSiteHeader nav a[href]',
      '#spSiteHeader [role="navigation"] a[href]',
      '[role="navigation"] a[href]'
    ].join(',');

    const anchors = Array.from(
      document.querySelectorAll(selectors)
    ) as HTMLAnchorElement[];

    const items: ITopNavItem[] = [];
    const seen = new Set<string>();

    anchors.forEach((anchor: HTMLAnchorElement) => {
      if (anchor.closest('[data-ami-top-banner="true"]')) {
        return;
      }

      const title = (anchor.innerText || anchor.textContent || '').trim();
      const url = anchor.href;

      if (!title || !url) {
        return;
      }

      // Skip icon-only / SharePoint system links
      if (title === '...' || title.length < 2) {
        return;
      }

      const key = `${title}|${url}`;

      if (seen.has(key)) {
        return;
      }

      seen.add(key);

      items.push({
        title,
        url
      });
    });

    return items;
  }

  private _hideOriginalSharePointNavigation(): void {
    const styleId = 'ami-hide-original-sharepoint-navigation';

    if (document.getElementById(styleId)) {
      return;
    }

    const style = document.createElement('style');
    style.id = styleId;

    style.innerHTML = `
      [data-automationid="HorizontalNav"],
      [data-automation-id="HorizontalNav"],
      [data-automationid="SiteNavigation"],
      [data-automation-id="SiteNavigation"],
      div:has(> [data-automationid="HorizontalNav"]),
      div:has(> [data-automation-id="HorizontalNav"]),
      div:has(> [data-automationid="SiteNavigation"]),
      div:has(> [data-automation-id="SiteNavigation"]) {
        display: none !important;
        height: 0 !important;
        min-height: 0 !important;
        max-height: 0 !important;
        padding: 0 !important;
        margin: 0 !important;
        border: 0 !important;
        overflow: hidden !important;
        visibility: hidden !important;
      }

      #spSiteHeader nav,
      #spSiteHeader [role="navigation"],
      #spSiteHeader div:has(nav),
      #spSiteHeader div:has([role="navigation"]) {
        display: none !important;
        height: 0 !important;
        min-height: 0 !important;
        max-height: 0 !important;
        padding: 0 !important;
        margin: 0 !important;
        border: 0 !important;
        overflow: hidden !important;
        visibility: hidden !important;
      }

      @media (max-width: 1200px) {
        #spSiteHeader,
        [data-automationid="SiteHeader"],
        [data-automation-id="SiteHeader"] {
          display: none !important;
          height: 0 !important;
          min-height: 0 !important;
          max-height: 0 !important;
          padding: 0 !important;
          margin: 0 !important;
          border: 0 !important;
          overflow: hidden !important;
          visibility: hidden !important;
        }
      }
    `;

    document.head.appendChild(style);
  }

  private _delay(milliseconds: number): Promise<void> {
    return new Promise((resolve: () => void) => {
      window.setTimeout(resolve, milliseconds);
    });
  }
  /*
  private _bindMoreMenuEvents(): void {
    const button = this._topPlaceholder?.domElement.querySelector(
      '[data-ami-more-button="true"]'
    ) as HTMLButtonElement | null;

    const menu = this._topPlaceholder?.domElement.querySelector(
      '[data-ami-more-menu="true"]'
    ) as HTMLElement | null;

    if (!button || !menu) {
      return;
    }

    button.addEventListener('click', (event: MouseEvent) => {
      event.preventDefault();
      event.stopPropagation();

      menu.classList.toggle(styles.moreMenuOpen);
    });

    document.addEventListener('click', () => {
      menu.classList.remove(styles.moreMenuOpen);
    });
  }
  */

  private _bindMoreMenuEvents(): void {
    const moreButton = this._topPlaceholder?.domElement.querySelector(
      '[data-ami-more-button="true"]'
    ) as HTMLButtonElement | null;

    const moreMenu = this._topPlaceholder?.domElement.querySelector(
      '[data-ami-more-menu="true"]'
    ) as HTMLElement | null;

    const hamburgerButton = this._topPlaceholder?.domElement.querySelector(
      '[data-ami-hamburger-button="true"]'
    ) as HTMLButtonElement | null;

    const mobileMenu = this._topPlaceholder?.domElement.querySelector(
      '[data-ami-mobile-menu="true"]'
    ) as HTMLElement | null;

    const closeMenus = (): void => {
      moreMenu?.classList.remove(styles.moreMenuOpen);
      mobileMenu?.classList.remove(styles.mobileMenuOpen);
    };

    if (moreButton && moreMenu) {
      moreButton.addEventListener('click', (event: MouseEvent) => {
        event.preventDefault();
        event.stopPropagation();

        mobileMenu?.classList.remove(styles.mobileMenuOpen);
        moreMenu.classList.toggle(styles.moreMenuOpen);
      });
    }

    if (hamburgerButton && mobileMenu) {
      hamburgerButton.addEventListener('click', (event: MouseEvent) => {
        event.preventDefault();
        event.stopPropagation();

        moreMenu?.classList.remove(styles.moreMenuOpen);
        mobileMenu.classList.toggle(styles.mobileMenuOpen);
      });
    }

    document.addEventListener('click', closeMenus);
  }

  private _normalizeNavigationUrl(url: string): string {
    if (!url) {
      return '#';
    }

    const cleanUrl = url.trim();

    if (cleanUrl === 'http://linkless.header/') {
      return '#';
    }

    const webAbsoluteUrl = this.context.pageContext.web.absoluteUrl.replace(/\/$/, '');
    const webRelativeUrl = this.context.pageContext.web.serverRelativeUrl.replace(/\/$/, '');
    const tenantOrigin = new URL(webAbsoluteUrl).origin;

    let finalUrl = '';

    if (cleanUrl.startsWith('http')) {
      finalUrl = cleanUrl;
    } else if (cleanUrl.startsWith('/')) {
      finalUrl = `${tenantOrigin}${cleanUrl}`;
    } else {
      finalUrl = `${webAbsoluteUrl}/${cleanUrl}`;
    }

    const duplicatePath = `${webRelativeUrl}${webRelativeUrl}`;

    while (finalUrl.indexOf(duplicatePath) !== -1) {
      finalUrl = finalUrl.replace(duplicatePath, webRelativeUrl);
    }

    return finalUrl;
  }

  private _setupCommandBarShortcut(): void {
    this._ensureCommandBarToggleStyle();
    this._applyCommandBarVisibility();

    if (!this._isCommandBarShortcutBound) {
      this._isCommandBarShortcutBound = true;

      document.addEventListener('keydown', (event: KeyboardEvent) => {
        const target = event.target as HTMLElement | null;

        const isTypingInsideInput =
          target?.tagName === 'INPUT' ||
          target?.tagName === 'TEXTAREA' ||
          target?.isContentEditable;

        if (isTypingInsideInput) {
          return;
        }

        if (event.shiftKey && event.code === 'KeyP') {
          event.preventDefault();
          event.stopPropagation();

          this._isCommandBarVisible = !this._isCommandBarVisible;
          this._applyCommandBarVisibility();
        }
      });
    }

    if (this._commandBarObserver) {
      this._commandBarObserver.disconnect();
    }

    this._commandBarObserver = new MutationObserver(() => {
      this._applyCommandBarVisibility();
    });

    this._commandBarObserver.observe(document.body, {
      childList: true,
      subtree: true
    });
  }

  private _ensureCommandBarToggleStyle(): void {
    const oldStyleIds = [
      'ami-hide-command-bar-wrapper',
      'ami-hide-command-bar-wrapper-v2'
    ];

    oldStyleIds.forEach((styleId: string) => {
      document.getElementById(styleId)?.remove();
    });

    const styleId = 'ami-command-bar-shortcut-style';

    if (document.getElementById(styleId)) {
      return;
    }

    const style = document.createElement('style');
    style.id = styleId;

    style.innerHTML = `
      body:not(.ami-command-bar-visible) .commandBarWrapper,
      body:not(.ami-command-bar-visible) .SPPageChrome-app .commandBarWrapper,
      body:not(.ami-command-bar-visible) div.commandBarWrapper,
      body:not(.ami-command-bar-visible) [class~="commandBarWrapper"] {
        display: none !important;
        height: 0 !important;
        min-height: 0 !important;
        max-height: 0 !important;
        padding: 0 !important;
        margin: 0 !important;
        border: 0 !important;
        overflow: hidden !important;
        visibility: hidden !important;
      }

      body.ami-command-bar-visible .commandBarWrapper,
      body.ami-command-bar-visible .SPPageChrome-app .commandBarWrapper,
      body.ami-command-bar-visible div.commandBarWrapper,
      body.ami-command-bar-visible [class~="commandBarWrapper"] {
        display: block !important;
        height: 48px !important;
        min-height: 48px !important;
        max-height: 48px !important;
        overflow: visible !important;
        visibility: visible !important;
      }
    `;

    document.head.appendChild(style);
  }

  private _applyCommandBarVisibility(): void {
    document.body.classList.toggle(
      'ami-command-bar-visible',
      this._isCommandBarVisible
    );

    const elements = document.querySelectorAll(
      '.commandBarWrapper, .SPPageChrome-app .commandBarWrapper, div.commandBarWrapper, [class~="commandBarWrapper"]'
    );

    elements.forEach((element: Element) => {
      const htmlElement = element as HTMLElement;

      if (this._isCommandBarVisible) {
        htmlElement.style.setProperty('display', 'block', 'important');
        htmlElement.style.setProperty('height', '48px', 'important');
        htmlElement.style.setProperty('min-height', '48px', 'important');
        htmlElement.style.setProperty('max-height', '48px', 'important');
        htmlElement.style.setProperty('overflow', 'visible', 'important');
        htmlElement.style.setProperty('visibility', 'visible', 'important');

        htmlElement.style.removeProperty('padding');
        htmlElement.style.removeProperty('margin');
        htmlElement.style.removeProperty('border');
      } else {
        htmlElement.style.setProperty('display', 'none', 'important');
        htmlElement.style.setProperty('height', '0', 'important');
        htmlElement.style.setProperty('min-height', '0', 'important');
        htmlElement.style.setProperty('max-height', '0', 'important');
        htmlElement.style.setProperty('padding', '0', 'important');
        htmlElement.style.setProperty('margin', '0', 'important');
        htmlElement.style.setProperty('border', '0', 'important');
        htmlElement.style.setProperty('overflow', 'hidden', 'important');
        htmlElement.style.setProperty('visibility', 'hidden', 'important');
      }
    });
  }

  private _escapeHtml(value: string): string {
    if (!value) {
      return '';
    }

    return value
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }

  private _onDispose(): void {
    console.log('AmiTopNav disposed.');
  }
}
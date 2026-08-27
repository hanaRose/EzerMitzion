 import { override } from '@microsoft/decorators';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';

import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';
import styles from './FooterApplicationCustomizer.module.scss';

export interface IFooterApplicationCustomizerProperties {
  listName?: string;
}

interface IFooterLink {
  Title: string;
  PageUrl: string;
  OpenInNewTab: boolean;
  Order: number;
  IsActive: boolean;
}

export default class FooterApplicationCustomizer
  extends BaseApplicationCustomizer<IFooterApplicationCustomizerProperties> {

  private _bottomPlaceholder: PlaceholderContent | undefined;
  private _footerLinks: IFooterLink[] = [];
  private _listName: string = 'FooterLinks';
  private _observer: MutationObserver | null = null;
  private _footerRendered: any;

  public async onInit(): Promise<void> {
    console.log('=== FOOTER INIT START ===');
    console.log('innerWidth:', window.innerWidth);
    console.log('userAgent:', navigator.userAgent);

    this._listName = this.properties.listName || 'FooterLinks';
    await this._ensureListExists();
    this._loadFontAwesome();

    const style = document.createElement('style');
    style.innerHTML += `
  #custom-spfx-footer,
  #custom-spfx-footer * {
    position: relative !important;
    bottom: auto !important;
    top: auto !important;
    left: auto !important;
    right: auto !important;
  }

  [class*="scrollableContent"] #custom-spfx-footer,
  [class*="ms-ScrollablePane"] #custom-spfx-footer {
    position: relative !important;
  }
`;
    this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceHolders);
    this._renderPlaceHolders();

    setTimeout(() => {
      console.log('=== MOBILE DIAGNOSTIC (3s after init) ===');
      console.log('_footerRendered:', this._footerRendered);
      console.log('custom-spfx-footer in DOM:', !!document.getElementById('custom-spfx-footer'));
      console.log('innerWidth:', window.innerWidth);

      const footer = document.getElementById('custom-spfx-footer');
      if (footer) {
        const rect = footer.getBoundingClientRect();
        const computed = window.getComputedStyle(footer);
        console.log('footer bounding rect:', JSON.stringify(rect));
        console.log('footer display:', computed.display);
        console.log('footer visibility:', computed.visibility);
        console.log('footer height:', computed.height);
        console.log('footer opacity:', computed.opacity);
        console.log('footer parent:', footer.parentElement?.id, footer.parentElement?.className?.substring(0, 80));
        console.log('footer innerHTML length:', footer.innerHTML.length);
      }

      const children = Array.from(document.body.children);
      console.log('Total body children:', children.length);
      children.slice(-5).forEach((el, i) => {
        const e = el as HTMLElement;
        const computed2 = window.getComputedStyle(e);
        console.log(`body[-${5 - i}]:`, e.tagName, e.id, e.className.substring(0, 60),
          'display:', computed2.display, 'height:', computed2.height, 'overflow:', computed2.overflow);
      });
    }, 3000);

    return Promise.resolve();
  }

  private _loadFontAwesome(): void {
    if (!document.querySelector('link[href*="font-awesome"]')) {
      const fontAwesomeLink = document.createElement('link');
      fontAwesomeLink.rel = 'stylesheet';
      fontAwesomeLink.href = 'https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css';
      fontAwesomeLink.crossOrigin = 'anonymous';
      document.head.appendChild(fontAwesomeLink);
      console.log('Font Awesome loaded');
    }
  }

  private async _ensureListExists(): Promise<void> {
    try {
      console.log(`Checking if list '${this._listName}' exists...`);

      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this._listName}')`,
        SPHttpClient.configurations.v1
      );

      if (response.ok) {
        console.log(`List '${this._listName}' already exists.`);
        return;
      }

      console.log(`List '${this._listName}' not found. Creating...`);
      await this._createList();

    } catch (error) {
      console.error('Error checking list existence:', error);
      await this._createList();
    }
  }

  private async _createList(): Promise<void> {
    console.log(`Creating list '${this._listName}'...`);

    const webUrl = this.context.pageContext.web.absoluteUrl;

    await this.context.spHttpClient.post(
      `${webUrl}/_api/web/lists`,
      SPHttpClient.configurations.v1,
      {
        headers: { 'Content-Type': 'application/json;odata=nometadata' },
        body: JSON.stringify({
          Title: this._listName,
          BaseTemplate: 100,
          Description: 'Footer Links data source'
        })
      }
    );

    console.log(`List '${this._listName}' created.`);

    const columns = [
      { FieldTypeKind: 2, Title: 'PageUrl' },
      { FieldTypeKind: 8, Title: 'IsActive' },
      { FieldTypeKind: 8, Title: 'OpenInNewTab' },
      { FieldTypeKind: 9, Title: 'Order' }
    ];

    for (const col of columns) {
      await this.context.spHttpClient.post(
        `${webUrl}/_api/web/lists/GetByTitle('${this._listName}')/fields`,
        SPHttpClient.configurations.v1,
        {
          headers: { 'Content-Type': 'application/json;odata=nometadata' },
          body: JSON.stringify(col)
        }
      );
      console.log(`Field '${col.Title}' added.`);
    }
  }

  private async _loadFooterLinks(): Promise<void> {
    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this._listName}')/items?$filter=IsActive eq 1&$orderby=Order&$select=Title,PageUrl,OpenInNewTab,Order,IsActive`,
        SPHttpClient.configurations.v1
      );

      if (response.ok) {
        const data = await response.json();
        this._footerLinks = data.value;
        console.log('Footer links loaded:', this._footerLinks);
      } else {
        console.warn('FooterLinks list not found or no active links available');
      }
    } catch (error) {
      console.error('Error loading footer links:', error);
    }
  }

  private _renderPlaceHolders(): void {
    if (this._footerRendered) {
      console.log('Footer already rendered, skipping');
      return;
    }

    if (document.getElementById('custom-spfx-footer')) {
      console.log('Footer div already in DOM, skipping');
      this._footerRendered = true;
      return;
    }

    console.log('_renderPlaceHolders called, innerWidth:', window.innerWidth);

    const megaFooter = document.querySelector('[class^="simpleFooterContainer"]') ||
      document.querySelector('[class*="simpleFooterContainer"]');

    console.log("megaFooter", megaFooter);
    if (megaFooter) {
      console.log('megaFooter found immediately');
      this._footerRendered = true;
      this._replaceMegaFooter(megaFooter as HTMLElement);
      return;
    }

    const isMobile = window.innerWidth <= 768;

    if (isMobile) {
      console.log('Mobile: appending footer to body');
      this._footerRendered = true;
      const footerDiv = document.createElement('div');
      footerDiv.id = 'custom-spfx-footer';
      footerDiv.style.cssText = 'position:relative;width:100%;height:auto;display:block;clear:both;box-sizing:border-box;';
      document.body.appendChild(footerDiv);
      this._replaceMegaFooter(footerDiv);
      return;
    }

    console.log('Desktop: starting observer + timeout fallback');
    this._observeForMegaFooter();
  }

  private _observeForMegaFooter(): void {
    console.log("_observeForMegaFooter");
    if (this._observer) {
      this._observer.disconnect();
    }

    this._observer = new MutationObserver((mutations) => {
      const megaFooter = document.querySelector('[class^="simpleFooterContainer"]') ||
        document.querySelector('[class*="simpleFooterContainer"]');

      if (megaFooter) {
        console.log("megafooter detected by observer", megaFooter);
        this._replaceMegaFooter(megaFooter as HTMLElement);

        if (this._observer) {
          this._observer.disconnect();
          this._observer = null;
        }
      }
    });

    this._observer.observe(document.body, {
      childList: true,
      subtree: true
    });

    console.log("MutationObserver started watching for megaFooter");
  }

  private _replaceMegaFooter(megaFooter: HTMLElement): void {
    console.log("_replaceMegaFooter", megaFooter);
    megaFooter.innerHTML = `
    <div class="${styles.footer}">  
    </div>
      <div class="${styles.footerContainer}">

        <!-- Right: Brand + Contact Info -->
        <div class="${styles.footerRight}">
          <div class="${styles.brandBlock}">
            <img src='https://ezermizionil.sharepoint.com/sites/portal/SiteAssets/Footer%2FEMFooterImage%2EJPG' />
          </div>
          <div class="${styles.contactBlock}">
            <div class="${styles.contactItem}">
              <img class="${styles.detailIcon}" src='https://ezermizionil.sharepoint.com/sites/portal/SiteAssets/Footer/mapIcon.JPG'>
              <span>הרב רבינוב 5, בני ברק</span>
            </div>
            <div class="${styles.contactItem}">
              <img class="${styles.detailIcon}" src='https://ezermizionil.sharepoint.com/sites/portal/SiteAssets/Footer/phoneIcon.JPG'>
              <span>03-6144444</span>
            </div>
            <div class="${styles.contactItem}">
              <img class="${styles.detailIcon}" src='https://ezermizionil.sharepoint.com/sites/portal/SiteAssets/Footer/donationIcon.JPG'>
              <span>לתרומות 1-800-236-236</span>
            </div>
            <div class="${styles.contactItem}">
              <img class="${styles.detailIcon}" src='https://ezermizionil.sharepoint.com/sites/portal/SiteAssets/Footer/faxIcon.JPG'>
              <span>03-6144445</span>
            </div>
          </div>
        </div>

        <!-- Left: Newsletter Signup -->
        <div class="${styles.footerLeft}">
          <p class="${styles.newsletterTitle}">הירשמו לניוזלטר שלנו וקבלו עידכונים</p>
          <div class="${styles.newsletterForm}">
            <input type="email" id="newsletter-input" placeholder="אימייל" class="${styles.newsletterInput}" />
          </div>
          <button id="newsletter-btn" class="${styles.newsletterButton}">שליחה</button>
        </div>

      </div>
    </div>
  `;

    const btn = document.getElementById('newsletter-btn');
    if (btn) {
      btn.addEventListener('click', () => this._registerNewsletter());
    }

    console.log("megaFooter content replaced successfully");
  }

  private async _registerNewsletter(): Promise<void> {
    console.log("_registerNewsletter");
    const valElem = document.getElementById('newsletter-input') as HTMLInputElement;
    const email = valElem.value;
    console.log("email", email);
    const button = document.getElementById('newsletter-btn') as HTMLButtonElement;
    const input = document.getElementById('newsletter-input') as HTMLInputElement;

    if (!email || !email.trim()) {
      input.style.borderColor = '#8b2020';
      return;
    }

    // ✅ בדיקת תקינות אימייל
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(email.trim())) {
      input.style.borderColor = '#8b2020';
      input.placeholder = 'אימייל לא תקין';
      input.value = '';
      setTimeout(() => {
        input.style.borderColor = '';
        input.placeholder = 'אימייל';
      }, 3000);
      return;
    }

    button.disabled = true;
    button.textContent = '...';

    try {
      const response = await this.context.spHttpClient.post(
        `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('NewsletterRegistration')/items`,
        SPHttpClient.configurations.v1,
        {
          headers: { 'Content-Type': 'application/json;odata=nometadata' },
          body: JSON.stringify({ Title: email.trim() })
        }
      );

      if (response.ok) {
        button.textContent = '✓';
        input.value = '';
        input.placeholder = 'תודה!';
        setTimeout(() => {
          input.placeholder = 'אימייל';
        }, 3000);
      } else {
        button.textContent = '✗';
        console.error('Newsletter registration failed:', response.statusText);
      }
    } catch (error) {
      button.textContent = '✗';
      console.error('Newsletter registration error:', error);
    }

    setTimeout(() => {
      button.disabled = false;
      button.textContent = 'שליחה';
    }, 3000);
  }

  private _generateFooterLinks(): string {
    if (this._footerLinks.length === 0) {
      return '';
    }

    return this._footerLinks.map((link, index) => {
      const url = link.PageUrl || '#';
      const target = link.OpenInNewTab ? 'target="_blank" rel="noopener noreferrer"' : '';
      const separator = index < this._footerLinks.length - 1
        ? `<span class="${styles.separator}">|</span>`
        : '';

      return `<a href="${url}" ${target} class="${styles.footerLink}">${link.Title}</a>${separator}`;
    }).join('');
  }

  private _onDispose(): void {
    console.log('Footer disposed');

    if (this._observer) {
      this._observer.disconnect();
      this._observer = null;
    }
  }
}
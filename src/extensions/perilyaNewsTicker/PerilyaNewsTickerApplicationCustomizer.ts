import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderName,
  PlaceholderContent
} from '@microsoft/sp-application-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

import * as strings from 'PerilyaNewsTickerApplicationCustomizerStrings';

const LOG_SOURCE: string = 'PerilyaNewsTickerApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IPerilyaNewsTickerApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class PerilyaNewsTickerApplicationCustomizer
  extends BaseApplicationCustomizer<IPerilyaNewsTickerApplicationCustomizerProperties> {

  private _topPlaceholder: PlaceholderContent | undefined;

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    // Create the top placeholder if available
    this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Top);
    if (this._topPlaceholder) {
      this._topPlaceholder.domElement.innerHTML = `<div id="perilya-news-ticker-container"></div>`;
    } else {
      // fallback: append to body
      const container = document.createElement('div');
      container.id = 'perilya-news-ticker-container';
      document.body.insertBefore(container, document.body.firstChild);
    }

    this._ensureStyles();
    this._renderNewsTicker();

    return Promise.resolve();
  }

  private _ensureStyles(): void {
    if (document.getElementById('perilya-news-ticker-styles')) {
      return;
    }

    const style = document.createElement('style');
    style.id = 'perilya-news-ticker-styles';
    style.innerHTML = `
    #perilya-news-ticker {
      position: fixed;
      top: 0;
      left: 0;
      width: 100%;
      background: #0b3b66;
      color: #ffffff;
      z-index: 99999;
      overflow: hidden;
      box-sizing: border-box;
      padding: 8px 0;
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial;
      font-size: 14px;
    }
    .test{
      color: red;}
    #perilya-news-ticker .ticker-inner {
      display: inline-block;
      white-space: nowrap;
      padding-left: 100%;
      will-change: transform;
      animation: perilya-ticker-scroll 36s linear infinite;
    }
    #perilya-news-ticker .ticker-item { display: inline-block; margin-right: 40px; }
    @keyframes perilya-ticker-scroll {
      0% { transform: translateX(0%); }
      100% { transform: translateX(-100%); }
    }
    `;
    document.head.appendChild(style);
  }

  private _renderNewsTicker(): void {
    const container = document.getElementById('perilya-news-ticker-container');
    if (!container) { return; }

    const ticker = document.createElement('div');
    ticker.id = 'perilya-news-ticker';
    ticker.setAttribute('role', 'region');
    ticker.setAttribute('aria-label', 'News ticker');
    ticker.innerHTML = `<div class="ticker-inner">Loading news...</div>`;
    container.appendChild(ticker);

    // make sure page content is not hidden under the fixed ticker
    const height = ticker.getBoundingClientRect().height || 36;
    const currentPadding = parseInt(window.getComputedStyle(document.body).paddingTop || '0', 10) || 0;
    document.body.style.paddingTop = (currentPadding + height) + 'px';

    this._fetchNewsItems()
      .then(titles => {
        if (!titles || titles.length === 0) {
          ticker.querySelector('.ticker-inner')!.textContent = 'No news items found.';
          return;
        }

        // Render items once (no duplication)
        const content = titles.map(t => `<span class="ticker-item">${this._escapeHtml(t)}</span>`).join('');
        (ticker.querySelector('.ticker-inner') as HTMLElement).innerHTML = content;
      })
      .catch(err => {
        console.error('Error loading news ticker items', err);
        ticker.querySelector('.ticker-inner')!.textContent = 'Failed to load news.';
      });
  }

  private _fetchNewsItems(): Promise<string[]> {
    const webUrl = this.context.pageContext.web.absoluteUrl;
    const requestUrl = `${webUrl}/_api/web/lists/GetByTitle('NewsTicker')/items?$select=Title,Id&$orderby=Id asc&$top=10`;

    return this.context.spHttpClient.get(requestUrl, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => response.json())
      .then(json => {
        if (!json || !json.value) { return []; }
        return json.value.map((i: any) => i.Title).filter((t: any) => !!t);
      });
  }

  private _escapeHtml(text: string): string {
    const div = document.createElement('div');
    div.appendChild(document.createTextNode(text));
    return div.innerHTML;
  }
}

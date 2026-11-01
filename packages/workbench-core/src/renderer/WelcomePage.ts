import type { RecentFolder } from '../contracts';
import type { EasyViewDesktopApi } from '../host';
import { tDesktop } from './desktopRenderer.i18n';

const COLLAPSED_RECENT_COUNT = 6;
const isMac = navigator.platform.toLowerCase().includes('mac');
const shortcut = (key: string, shift = false) => (isMac ? `${shift ? '⇧' : ''}⌘${key}` : `Ctrl+${shift ? 'Shift+' : ''}${key}`);

const ICONS = {
  logo: '<svg viewBox="0 0 16 16" aria-hidden="true"><path d="M6.4 4.5C5.6 3.1 4.5 2.5 3.3 2.5C2.1 2.5 1.5 3.2 1.5 4.4V11.2C1.5 12.4 2.1 13 3.3 13H4.5C5.3 13 5.9 13.3 6.4 13.8C7 13.3 7.7 13 8.6 13H12.8C13.9 13 14.5 12.4 14.5 11.3V6L11.3 2.8H8.4C7.1 2.8 6.4 3.4 6.4 4.5V13.5"/><path d="M11.3 2.8V5C11.3 5.7 11.6 6 12.3 6H14.5M8.3 8.3H12.1M8.3 10.4H11.1"/></svg>',
  folder: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M3.5 7.5a2 2 0 0 1 2-2h4l2 2.5h7a2 2 0 0 1 2 2v7.5a2 2 0 0 1-2 2h-13a2 2 0 0 1-2-2Z"/></svg>',
  file: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M14 3.5H7.5a2 2 0 0 0-2 2v13a2 2 0 0 0 2 2h9a2 2 0 0 0 2-2V8Z"/><path d="M14 3.5V8h4.5M9 13h6M9 16.5h4"/></svg>',
  window: '<svg viewBox="0 0 24 24" aria-hidden="true"><rect x="3.5" y="5" width="17" height="14" rx="2"/><path d="M3.5 9h17M12 12v4M10 14h4"/></svg>',
  resume: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 12a8 8 0 1 0 2.4-5.7"/><path d="M4 4.5v4h4"/><path d="M10.5 9.2v5.6l4.3-2.8Z"/></svg>',
  close: '<svg viewBox="0 0 16 16" aria-hidden="true"><path d="M4.5 4.5l7 7M11.5 4.5l-7 7"/></svg>',
} as const;

export interface WelcomePageOptions {
  api: EasyViewDesktopApi;
  container: HTMLElement;
  onOpenFolder(): void;
  onOpenRecentFolder(folderPath: string): void;
  onResume(): void;
  onShowWelcome(): void;
}

/** `empty` is the light hint shown when a folder is open but no file is. */
export type WelcomePageMode = 'welcome' | 'empty';

export interface WelcomeResumeInfo {
  folderPath: string | null;
  tabCount: number;
  activeFileName: string | null;
}

/** Start screen shown at launch (over restored tabs) and whenever no tab is open. */
export class WelcomePage {
  private visible = false;
  private expanded = false;
  private currentRootPath: string | null = null;
  private resume: WelcomeResumeInfo | null = null;
  private mode: WelcomePageMode = 'welcome';
  private folders: RecentFolder[] = [];
  private loadToken = 0;

  constructor(private readonly options: WelcomePageOptions) {
    options.container.classList.add('desktop-welcome');
    options.container.dataset.mode = this.mode;
    options.container.hidden = true;
  }

  setVisible(visible: boolean): void {
    if (this.visible === visible) return;
    this.visible = visible;
    this.options.container.hidden = !visible;
    if (visible) void this.refresh();
  }

  setCurrentRoot(rootPath: string | null): void {
    if (this.currentRootPath === rootPath) return;
    this.currentRootPath = rootPath;
    if (this.visible) void this.refresh();
  }

  isVisible(): boolean {
    return this.visible;
  }

  getMode(): WelcomePageMode {
    return this.mode;
  }

  setMode(mode: WelcomePageMode): void {
    if (this.mode === mode) return;
    this.mode = mode;
    this.options.container.dataset.mode = mode;
    if (this.visible) void this.refresh();
  }

  setResume(resume: WelcomeResumeInfo | null): void {
    const unchanged = this.resume?.tabCount === resume?.tabCount && this.resume?.activeFileName === resume?.activeFileName
      && this.resume?.folderPath === resume?.folderPath;
    this.resume = resume;
    if (this.visible && !unchanged) this.render();
  }

  async refresh(): Promise<void> {
    const token = ++this.loadToken;
    const result = await this.options.api.workspace.getRecentFolders();
    if (token !== this.loadToken) return;
    this.folders = result.ok ? result.value : [];
    this.render();
  }

  private render(): void {
    const { container } = this.options;
    container.replaceChildren();
    if (this.mode === 'empty') {
      container.appendChild(this.renderEmptyHint());
      return;
    }
    const inner = element('div', 'desktop-welcome-inner');

    const brand = element('div', 'desktop-welcome-brand');
    const logo = element('span', 'desktop-welcome-logo');
    logo.innerHTML = ICONS.logo;
    const title = element('h1', 'desktop-welcome-title', '文档工作台');
    brand.append(logo, title);
    const subtitle = element('p', 'desktop-welcome-subtitle', tDesktop('welcome.subtitle'));

    const actions = element('div', 'desktop-welcome-actions');
    actions.append(
      this.actionCard(ICONS.folder, tDesktop('welcome.openFolder'), shortcut('O', true), () => this.options.onOpenFolder(), 'open-folder'),
      this.actionCard(ICONS.file, tDesktop('welcome.openFile'), shortcut('O'), () => {
        void this.options.api.document.open().then((result) => { if (result.ok && result.value) this.options.onResume(); });
      }, 'open-file'),
    );
    if (this.options.api.capabilities?.newWindow !== false) {
      actions.appendChild(this.actionCard(ICONS.window, tDesktop('welcome.newWindow'), '', () => { void this.options.api.app.newWindow(); }, 'new-window'));
    }

    inner.append(brand, subtitle);
    if (this.resume) inner.appendChild(this.renderResume(this.resume));
    inner.append(actions, this.renderRecent());
    container.appendChild(inner);
  }

  private renderEmptyHint(): HTMLElement {
    const hint = element('div', 'desktop-welcome-hint');
    const logo = element('span', 'desktop-welcome-hint-logo');
    logo.innerHTML = ICONS.logo;
    hint.append(logo, element('p', 'desktop-welcome-hint-text', tDesktop('welcome.emptyHint')));

    const shortcuts = element('dl', 'desktop-welcome-hint-shortcuts');
    for (const [label, keys] of [[tDesktop('welcome.openFile'), shortcut('O')], [tDesktop('welcome.openFolder'), shortcut('O', true)]] as const) {
      shortcuts.append(element('dt', '', label), element('dd', '', keys));
    }
    hint.appendChild(shortcuts);

    const more = element('button', 'desktop-welcome-link', tDesktop('welcome.showWelcome')) as HTMLButtonElement;
    more.type = 'button';
    more.dataset.action = 'show-welcome';
    more.addEventListener('click', () => this.options.onShowWelcome());
    hint.appendChild(more);
    return hint;
  }

  private renderResume(resume: WelcomeResumeInfo): HTMLButtonElement {
    const button = element('button', 'desktop-welcome-resume') as HTMLButtonElement;
    button.type = 'button';
    button.dataset.action = 'resume';
    const icon = element('span', 'desktop-welcome-action-icon');
    icon.innerHTML = ICONS.resume;
    const text = element('span', 'desktop-welcome-resume-text');
    text.appendChild(element('span', 'desktop-welcome-action-label', tDesktop('welcome.resume')));
    const folderName = resume.folderPath?.split(/[\\/]/).filter(Boolean).at(-1);
    const detail = [folderName, resume.tabCount === 1 ? tDesktop('welcome.tabCountOne') : tDesktop('welcome.tabCount', { count: resume.tabCount }), resume.activeFileName].filter(Boolean).join(' · ');
    text.appendChild(element('span', 'desktop-welcome-resume-detail', detail));
    button.append(icon, text, element('span', 'desktop-welcome-action-hint', 'Esc'));
    button.addEventListener('click', () => this.options.onResume());
    return button;
  }

  private actionCard(icon: string, label: string, hint: string, onClick: () => void, id: string): HTMLButtonElement {
    const button = element('button', 'desktop-welcome-action') as HTMLButtonElement;
    button.type = 'button';
    button.dataset.action = id;
    const iconElement = element('span', 'desktop-welcome-action-icon');
    iconElement.innerHTML = icon;
    button.append(iconElement, element('span', 'desktop-welcome-action-label', label));
    if (hint) button.appendChild(element('span', 'desktop-welcome-action-hint', hint));
    button.addEventListener('click', onClick);
    return button;
  }

  private renderRecent(): HTMLElement {
    const section = element('section', 'desktop-welcome-recent');
    const header = element('div', 'desktop-welcome-recent-header');
    header.appendChild(element('span', 'desktop-welcome-recent-title', tDesktop('welcome.recentTitle')));
    if (this.folders.length > COLLAPSED_RECENT_COUNT) {
      const toggle = element('button', 'desktop-welcome-link', this.expanded ? tDesktop('welcome.collapse') : tDesktop('welcome.showAll', { count: this.folders.length })) as HTMLButtonElement;
      toggle.type = 'button';
      toggle.addEventListener('click', () => { this.expanded = !this.expanded; this.render(); });
      header.appendChild(toggle);
    }
    section.appendChild(header);

    if (this.folders.length === 0) {
      section.appendChild(element('p', 'desktop-welcome-empty', tDesktop('welcome.recentEmpty')));
      return section;
    }

    const list = element('ul', 'desktop-welcome-recent-list');
    const shown = this.expanded ? this.folders : this.folders.slice(0, COLLAPSED_RECENT_COUNT);
    for (const folder of shown) list.appendChild(this.renderRecentItem(folder));
    section.appendChild(list);
    return section;
  }

  private renderRecentItem(folder: RecentFolder): HTMLElement {
    const item = element('li', 'desktop-welcome-recent-item');
    item.classList.toggle('is-missing', !folder.exists);
    item.classList.toggle('is-current', folder.path === this.currentRootPath);

    const open = element('button', 'desktop-welcome-recent-open') as HTMLButtonElement;
    open.type = 'button';
    open.title = folder.exists ? folder.path : tDesktop('welcome.missingPath', { path: folder.path });
    open.disabled = !folder.exists;
    const name = element('span', 'desktop-welcome-recent-name');
    name.appendChild(element('span', 'desktop-welcome-recent-label', folder.name));
    if (folder.path === this.currentRootPath) name.appendChild(element('span', 'desktop-welcome-badge', tDesktop('welcome.current')));
    if (!folder.exists) name.appendChild(element('span', 'desktop-welcome-badge is-warning', tDesktop('welcome.missing')));
    open.append(name, element('span', 'desktop-welcome-recent-path', folder.displayPath));
    open.addEventListener('click', () => this.options.onOpenRecentFolder(folder.path));

    const remove = element('button', 'desktop-welcome-recent-remove') as HTMLButtonElement;
    remove.type = 'button';
    remove.title = tDesktop('welcome.removeRecent');
    remove.setAttribute('aria-label', tDesktop('welcome.removeRecentNamed', { name: folder.name }));
    remove.innerHTML = ICONS.close;
    remove.addEventListener('click', async (event) => {
      event.stopPropagation();
      const result = await this.options.api.workspace.removeRecentFolder(folder.path);
      if (result.ok) await this.refresh();
    });

    item.append(open, remove);
    return item;
  }
}

function element(tag: string, className: string, text?: string): HTMLElement {
  const node = document.createElement(tag);
  node.className = className;
  if (text !== undefined) node.textContent = text;
  return node;
}

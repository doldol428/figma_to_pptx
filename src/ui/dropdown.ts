/**
 * 커스텀 드롭다운.
 *
 * 네이티브 `<select>` 의 팝업 목록은 브라우저가 아니라 OS 가 그린다. 모서리 반경도,
 * 배경색도, Figma 테마 변수도 닿지 않는다. 그래서 트리거 버튼 + 절대배치 목록으로 직접 만든다.
 *
 * 목록은 `position: absolute` 라 body 높이에 잡히지 않는다. 열고 닫을 때마다
 * `onResize` 를 불러 플러그인 창 높이를 다시 계산해야 iframe 에 잘리지 않는다.
 */

export interface DropdownOption {
  value: number;
  /** 닫혀 있을 때 트리거에 보이는 문구 */
  label: string;
  /** 목록 행 문구 (없으면 label) */
  rowLabel?: string;
  /** 행 오른쪽 보조 문구 */
  note?: string;
}

export class Dropdown {
  private isOpen = false;
  private selectedIndex: number;
  private activeIndex: number;
  private rows: HTMLElement[] = [];

  constructor(
    root: HTMLElement,
    private readonly trigger: HTMLButtonElement,
    private readonly menu: HTMLElement,
    private readonly options: DropdownOption[],
    initialValue: number,
    private readonly onResize: () => void,
  ) {
    const found = options.findIndex((o) => o.value === initialValue);
    this.selectedIndex = found < 0 ? 0 : found;
    this.activeIndex = this.selectedIndex;

    this.build();
    this.syncTrigger();

    trigger.addEventListener('click', () => this.toggle());
    trigger.addEventListener('keydown', (e) => this.onKeyDown(e));
    // 바깥을 누르면 닫는다. click 이 아니라 mousedown 이라야 다른 컨트롤을 한 번에 누를 수 있다.
    document.addEventListener('mousedown', (e) => {
      if (this.isOpen && !root.contains(e.target as Node)) this.close();
    });
  }

  get value(): number {
    return this.options[this.selectedIndex].value;
  }

  setDisabled(disabled: boolean): void {
    this.trigger.disabled = disabled;
    if (disabled) this.close();
  }

  /** 열려 있으면 목록 아래 끝의 뷰포트 y 좌표, 닫혀 있으면 null */
  menuBottom(): number | null {
    return this.isOpen ? this.menu.getBoundingClientRect().bottom : null;
  }

  private build(): void {
    // 문구는 전부 코드 안의 상수라 이스케이프가 필요 없다.
    this.menu.innerHTML = this.options
      .map((o, i) => {
        const note = o.note ? `<span class="opt-note">${o.note}</span>` : '';
        return `<div class="opt" role="option" id="${this.menu.id}-${i}">${o.rowLabel ?? o.label}${note}</div>`;
      })
      .join('');

    this.rows = Array.from(this.menu.querySelectorAll<HTMLElement>('.opt'));
    this.rows.forEach((row, i) => {
      row.addEventListener('click', () => this.select(i));
      row.addEventListener('mouseenter', () => this.setActive(i));
    });
    this.syncRows();
  }

  private syncTrigger(): void {
    this.trigger.textContent = this.options[this.selectedIndex].label;
  }

  private syncRows(): void {
    this.rows.forEach((row, i) => {
      row.setAttribute('aria-selected', String(i === this.selectedIndex));
      row.classList.toggle('active', i === this.activeIndex);
    });
  }

  private setActive(i: number): void {
    this.activeIndex = (i + this.rows.length) % this.rows.length;
    this.trigger.setAttribute('aria-activedescendant', this.rows[this.activeIndex].id);
    this.syncRows();
  }

  private toggle(): void {
    if (this.isOpen) this.close();
    else this.open();
  }

  private open(): void {
    if (this.isOpen || this.trigger.disabled) return;
    this.isOpen = true;
    this.menu.classList.remove('hidden');
    this.trigger.setAttribute('aria-expanded', 'true');
    this.setActive(this.selectedIndex);
    this.onResize();
  }

  private close(): void {
    if (!this.isOpen) return;
    this.isOpen = false;
    this.menu.classList.add('hidden');
    this.trigger.setAttribute('aria-expanded', 'false');
    this.trigger.removeAttribute('aria-activedescendant');
    this.onResize();
  }

  private select(i: number): void {
    this.selectedIndex = i;
    this.activeIndex = i;
    this.syncTrigger();
    this.syncRows();
    this.close();
  }

  private onKeyDown(e: KeyboardEvent): void {
    switch (e.key) {
      case 'ArrowDown':
        e.preventDefault();
        if (this.isOpen) this.setActive(this.activeIndex + 1);
        else this.open();
        break;
      case 'ArrowUp':
        e.preventDefault();
        if (this.isOpen) this.setActive(this.activeIndex - 1);
        else this.open();
        break;
      case 'Home':
        if (this.isOpen) {
          e.preventDefault();
          this.setActive(0);
        }
        break;
      case 'End':
        if (this.isOpen) {
          e.preventDefault();
          this.setActive(this.rows.length - 1);
        }
        break;
      case 'Enter':
      case ' ':
        // preventDefault 로 버튼의 기본 click 을 막지 않으면 토글과 선택이 겹쳐 일어난다.
        e.preventDefault();
        if (this.isOpen) this.select(this.activeIndex);
        else this.open();
        break;
      case 'Escape':
        if (this.isOpen) {
          e.preventDefault();
          this.close();
        }
        break;
      case 'Tab':
        this.close();
        break;
      default:
        break;
    }
  }
}

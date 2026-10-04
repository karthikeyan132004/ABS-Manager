import * as React from 'react';
import { Icon } from '@fluentui/react';
import styles from './ManagerMaster.module.scss';

export interface INavItem {
  key: string;
  label: string;
  /** Fluent UI icon name. */
  icon: string;
  /** Shown as a pill on the right of the row. Omitted when undefined. */
  badge?: number;
  /** Groups rows under a heading. Rows sharing a group stay together. */
  group: string;
}

export interface ISideNavProps {
  items: INavItem[];
  selectedKey: string;
  onSelect: (key: string) => void;
  /** Brand text at the top of the rail. */
  brand: string;
  collapsed: boolean;
  onToggleCollapsed: () => void;
}

/**
 * The dark red rail. Groups render in the order their first item appears, so
 * the caller controls ordering by ordering the array.
 */
const SideNav: React.FunctionComponent<ISideNavProps> = (props: ISideNavProps) => {
  const { items, selectedKey, onSelect, brand, collapsed, onToggleCollapsed } = props;

  const groups: string[] = [];
  for (const item of items) {
    if (groups.indexOf(item.group) < 0) {
      groups.push(item.group);
    }
  }

  return (
    <nav className={`${styles.sideNav} ${collapsed ? styles.sideNavCollapsed : ''}`} aria-label="Dashboard sections">
      <div className={styles.brandRow}>
        <span className={styles.brandMark} aria-hidden="true">
          <Icon iconName="BIDashboard" />
        </span>
        {!collapsed && <span className={styles.brandText}>{brand}</span>}
        <button
          type="button"
          className={styles.railToggle}
          onClick={onToggleCollapsed}
          aria-label={collapsed ? 'Expand navigation' : 'Collapse navigation'}
          title={collapsed ? 'Expand navigation' : 'Collapse navigation'}
        >
          <Icon iconName={collapsed ? 'DoubleChevronRight' : 'DoubleChevronLeft'} />
        </button>
      </div>

      {groups.map((group: string) => (
        <div key={group} className={styles.navGroup}>
          {!collapsed && <div className={styles.navGroupLabel}>{group}</div>}
          {items
            .filter((item: INavItem): boolean => item.group === group)
            .map((item: INavItem) => {
              const active: boolean = item.key === selectedKey;
              return (
                <button
                  type="button"
                  key={item.key}
                  className={`${styles.navItem} ${active ? styles.navItemActive : ''}`}
                  onClick={(): void => onSelect(item.key)}
                  aria-current={active ? 'page' : undefined}
                  title={collapsed ? item.label : undefined}
                >
                  <Icon iconName={item.icon} className={styles.navIcon} />
                  {!collapsed && <span className={styles.navLabel}>{item.label}</span>}
                  {!collapsed && item.badge !== undefined && (
                    <span className={styles.navBadge}>{item.badge}</span>
                  )}
                </button>
              );
            })}
        </div>
      ))}
    </nav>
  );
};

export default SideNav;

import * as React from 'react';
import styles from './ManagerMaster.module.scss';
import { ALL_STATUSES, displayStatus } from '../../../statuses';

export interface IStageChip {
  status: string;
  count: number;
}

export interface IStageChipsProps {
  chips: IStageChip[];
  /** Currently filtered stage, if any. */
  selected?: string;
  onSelect: (status: string | undefined) => void;
}

/**
 * The stage rail. Thirteen full KPI cards was a wall of mostly-zero tiles, so
 * each stage is a compact chip instead: its pipeline number, its name, its
 * count, and a share bar showing how much of the board sits there.
 */
const StageChips: React.FunctionComponent<IStageChipsProps> = (props: IStageChipsProps) => {
  const { chips, selected, onSelect } = props;

  const busiest: number = chips.reduce((highest: number, chip: IStageChip): number =>
    Math.max(highest, chip.count), 0);

  return (
    <div className={styles.chipRail}>
      {chips.map((chip: IStageChip) => {
        const active: boolean = selected === chip.status;
        const position: number = ALL_STATUSES.indexOf(chip.status) + 1;
        const share: number = busiest === 0 ? 0 : Math.round((chip.count / busiest) * 100);

        return (
          <button
            type="button"
            key={chip.status}
            className={`${styles.chip} ${active ? styles.chipActive : ''} ${chip.count === 0 ? styles.chipEmpty : ''}`}
            onClick={(): void => onSelect(active ? undefined : chip.status)}
            aria-pressed={active}
            title={`${chip.status} - ${chip.count} quotation(s)`}
          >
            <span className={styles.chipTop}>
              <span className={styles.chipNumber}>{position > 0 ? (position < 10 ? `0${position}` : String(position)) : '--'}</span>
              <span className={styles.chipCount}>{chip.count.toLocaleString()}</span>
            </span>
            <span className={styles.chipLabel}>{displayStatus(chip.status)}</span>
            <span className={styles.chipTrack}>
              {chip.count > 0 && <span className={styles.chipBar} style={{ width: `${share}%` }} />}
            </span>
          </button>
        );
      })}
    </div>
  );
};

export default StageChips;

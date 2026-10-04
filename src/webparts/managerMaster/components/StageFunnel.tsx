import * as React from 'react';
import { Text } from '@fluentui/react';
import { displayStatus } from '../../../statuses';
import styles from './ManagerMaster.module.scss';

export interface IStageCount {
  status: string;
  count: number;
}

export interface IStageFunnelProps {
  title: string;
  stages: IStageCount[];
  /** Clicking a bar jumps to that stage. */
  onSelect?: (status: string) => void;
}

/**
 * One measure across ordered stages, so: horizontal bars in pipeline order,
 * never sorted by value. A single series, so one brand hue throughout and no
 * legend - length carries the magnitude and every bar is directly labelled.
 */
const StageFunnel: React.FunctionComponent<IStageFunnelProps> = (props: IStageFunnelProps) => {
  const { title, stages, onSelect } = props;

  const max: number = stages.reduce((highest: number, stage: IStageCount): number =>
    Math.max(highest, stage.count), 0);
  const total: number = stages.reduce((sum: number, stage: IStageCount): number => sum + stage.count, 0);

  return (
    <section className={styles.panelCard}>
      <header className={styles.panelCardHeader}>
        <Text variant="mediumPlus" className={styles.panelCardTitle}>{title}</Text>
        <Text variant="small" className={styles.panelCardMeta}>{total.toLocaleString()} in total</Text>
      </header>

      {total === 0 ? (
        <Text variant="small" className={styles.emptyNote}>
          Nothing to show yet. Either no items carry these statuses, or the status column has not been matched.
        </Text>
      ) : (
        <ol className={styles.funnel}>
          {stages.map((stage: IStageCount) => {
            // Width is relative to the busiest stage so short stages stay readable.
            const width: number = max === 0 ? 0 : Math.round((stage.count / max) * 100);
            const row: React.ReactElement = (
              <>
                <span className={styles.funnelLabel} title={stage.status}>{displayStatus(stage.status)}</span>
                <span className={styles.funnelTrack}>
                  {stage.count > 0 && (
                    <span className={styles.funnelBar} style={{ width: `${width}%` }} aria-hidden="true" />
                  )}
                </span>
                <span className={styles.funnelValue}>{stage.count.toLocaleString()}</span>
              </>
            );

            return (
              <li key={stage.status} className={styles.funnelRow}>
                {onSelect === undefined ? (
                  <div className={styles.funnelRowInner}>{row}</div>
                ) : (
                  <button
                    type="button"
                    className={`${styles.funnelRowInner} ${styles.funnelRowButton}`}
                    onClick={(): void => onSelect(stage.status)}
                    title={`${displayStatus(stage.status)}: ${stage.count}`}
                  >
                    {row}
                  </button>
                )}
              </li>
            );
          })}
        </ol>
      )}
    </section>
  );
};

export default StageFunnel;

import * as React from 'react';
import { Icon } from '@fluentui/react';
import styles from './ManagerMaster.module.scss';

export type KpiTone = 'neutral' | 'warning' | 'danger' | 'success' | 'brand';

export interface IKpiCardProps {
  label: string;
  value: number;
  iconName: string;
  tone?: KpiTone;
  /** Small caption under the value, e.g. which field the number came from. */
  hint?: string;
  /** Overrides the rendered figure, for money and other formatted values. */
  display?: string;
  /** Makes the card behave as a filter toggle. */
  onClick?: () => void;
  active?: boolean;
}

const TONE_CLASS: Record<KpiTone, string> = {
  neutral: styles.kpiNeutral,
  warning: styles.kpiWarning,
  danger: styles.kpiDanger,
  success: styles.kpiSuccess,
  brand: styles.kpiBrand
};

const KpiCard: React.FunctionComponent<IKpiCardProps> = (props: IKpiCardProps) => {
  const tone: KpiTone = props.tone ?? 'neutral';
  const interactive: boolean = props.onClick !== undefined;

  const className: string = [
    styles.kpiCard,
    TONE_CLASS[tone],
    interactive ? styles.kpiCardInteractive : '',
    props.active === true ? styles.kpiCardActive : ''
  ].join(' ');

  const body: React.ReactElement = (
    <>
      <span className={styles.kpiIconWrap} aria-hidden="true">
        <Icon iconName={props.iconName} className={styles.kpiIcon} />
      </span>
      <div className={styles.kpiBody}>
        <div className={styles.kpiValue}>{props.display ?? props.value.toLocaleString()}</div>
        <div className={styles.kpiLabel}>{props.label}</div>
        {props.hint !== undefined && props.hint !== '' && (
          <div className={styles.kpiHint}>{props.hint}</div>
        )}
      </div>
    </>
  );

  if (!interactive) {
    return <div className={className}>{body}</div>;
  }

  return (
    <button type="button" className={className} onClick={props.onClick} aria-pressed={props.active === true}>
      {body}
    </button>
  );
};

export default KpiCard;

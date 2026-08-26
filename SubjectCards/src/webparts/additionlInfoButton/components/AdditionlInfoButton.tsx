import * as React from 'react';
import styles from './AdditionlInfoButton.module.scss';
import type { IAdditionalInfoButtonProps } from './IAdditionlInfoButtonProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { IAdditionalInfoButtonItem, ListService } from '../services/ListService';

const AdditionalInfoButton: React.FC<IAdditionalInfoButtonProps> = ({ listService }) => {
  const [AdditionalInfoButton, setAdditionalInfoButton] = React.useState<IAdditionalInfoButtonItem[]>([]);
  const [loading, setLoading] = React.useState(true);
  const [error, setError] = React.useState<string | null>(null);

  React.useEffect(() => {
    listService.getAdditionalInfoButtonData()
      .then(setAdditionalInfoButton)
      .catch(() => setError('Failed to load Additional Info Button.'))
      .finally(() => setLoading(false));
  }, []);

  if (loading) return <div className={styles.status}>Loading Quick Links...</div>;
  if (error) return <div className={styles.status}>{error}</div>;
  if (!AdditionalInfoButton.length) return <div className={styles.status}>No Additional button item found. Add items to the <strong>AdditionalInfoButton</strong> list.</div>;

  const handleClick = (item: IAdditionalInfoButtonItem): void => {
    const url = `${item.Url}`;
    window.open(url, '_blank');
  };

  return (
    <div>
      {AdditionalInfoButton.map(item => (
        <div
          className={styles.additionalInfoBtn}
          onClick={() => handleClick(item)}
          role="button"
          tabIndex={0}
          onKeyDown={e => e.key === 'Enter' && handleClick(item)}
        >
          <span>{item.Title}</span>
          {/* Icon on the left */}
        </div>
      ))}
    </div>
  );
};

export default AdditionalInfoButton;
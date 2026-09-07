 import * as React from 'react';
import { IWidgetTwoProps } from './IWidgetTwoProps';
import { IHappyMomentsItem } from '../../services/HappyMomentsService';
import styles from './WidgetTwo.module.scss';

const eventIconMap: Record<string, string> = {
  'לידה': 'https://ezermizionil.sharepoint.com/sites/portal/SiteAssets/HappyMoments/baby.png',
  'חתונה': 'https://ezermizionil.sharepoint.com/sites/portal/SiteAssets/HappyMoments/flowers.png',
  'אירוסין': 'https://ezermizionil.sharepoint.com/sites/portal/SiteAssets/HappyMoments/flowers.png',
  'בר/בת מצווה': 'https://ezermizionil.sharepoint.com/sites/portal/SiteAssets/HappyMoments/flowers.png',
  'אחר': 'https://ezermizionil.sharepoint.com/sites/portal/SiteAssets/HappyMoments/flowers.png',
};

const WidgetTwo: React.FC<IWidgetTwoProps> = ({ listService }) => {
  const [items, setItems] = React.useState<IHappyMomentsItem[]>([]);
  const [loading, setLoading] = React.useState(true);
  const [error, setError] = React.useState<string | null>(null);

  React.useEffect(() => {
    listService.getItems()
      .then(setItems)
      .catch(() => setError('Failed to load items.'))
      .finally(() => setLoading(false));
  }, []);

  if (loading) return <div className={styles.status}>Loading...</div>;
  if (error) return <div className={styles.status}>{error}</div>;

  return (
    <div className={styles.wrapper} dir="rtl">
      <div className={styles.header}>
        <span className={styles.handwriting}>חוגגים ביחד</span>
        <h2 className={styles.sectionTitle}>חולקים רגעים של שמחה</h2>
        <div className={styles.smallUnderline} />
      </div>

      <div className={styles.list}>
        {items.map(item => {
          const iconUrl = eventIconMap[item.EventType] ?? eventIconMap['אחר'];

          return (
            <div key={item.Id} className={styles.card}>
              <div className={styles.cardContent}>
                <span className={styles.cardTitle}>{item.Title}</span>
                {item.SubTitle && (
                  <span className={styles.cardSubTitle}>{item.SubTitle}</span>
                )}
              </div>
              <img src={iconUrl} alt={item.EventType} className={styles.icon} />
            </div>
          );
        })}
      </div>
      <a
      className={styles.updateButton}
      href={`${window.location.origin}/sites/portal/Lists/HappyMomentsSubmissions/NewForm.aspx?Source=${encodeURIComponent(
        `${window.location.origin}/sites/portal`
      )}`}
    >
      עדכנו על שמחה בארגון
    </a>
    </div>
  );
};

export default WidgetTwo;
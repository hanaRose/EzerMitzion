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
          const openAttachment = (): void => {
            if (!item.AttachmentUrl) {
              return;
            }

            window.open(
              item.AttachmentUrl,
              '_blank',
              'noopener,noreferrer'
            );
          };

          return (
            <div
              key={item.Id}
              className={styles.card}
              role={item.AttachmentUrl ? 'link' : undefined}
              tabIndex={item.AttachmentUrl ? 0 : undefined}
              title={item.AttachmentUrl ? 'פתיחת ההזמנה בלשונית חדשה' : undefined}
              style={item.AttachmentUrl ? { cursor: 'pointer' } : undefined}
              onClick={item.AttachmentUrl ? openAttachment : undefined}
              onKeyDown={
                item.AttachmentUrl
                  ? (event: React.KeyboardEvent<HTMLDivElement>): void => {
                      if (event.key === 'Enter' || event.key === ' ') {
                        event.preventDefault();
                        openAttachment();
                      }
                    }
                  : undefined
              }
            >
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
      href="https://ezermizionil.sharepoint.com/:l:/s/portal/JABnni72sK96R4WADmp5YWQIAU3llqafRoCPia66vgOzmPE?nav=NWVlZTRjODItNWM3Ny00ZDdjLTlmZWUtYzlhYjJmNTRhOWI1"
    >
      עדכנו על שמחה בארגון
    </a>
    </div>
  );
};

export default WidgetTwo;
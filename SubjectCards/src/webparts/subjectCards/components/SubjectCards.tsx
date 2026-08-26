 import * as React from 'react';
import { ISubjectCardsProps } from './ISubjectCardsProps';
import { ICardItem } from '../services/ListService';
import styles from './SubjectCards.module.scss';

const Cards: React.FC<ISubjectCardsProps> = ({ listService }) => {
  const [cards, setCards] = React.useState<ICardItem[]>([]);
  const [loading, setLoading] = React.useState(true);
  const [error, setError] = React.useState<string | null>(null);

  React.useEffect(() => {
    listService.getCards()
      .then(setCards)
      .catch(() => setError('Failed to load cards.'))
      .finally(() => setLoading(false));
  }, []);

  if (loading) return <div className={styles.status}>Loading...</div>;
  if (error) return <div className={styles.status}>{error}</div>;
  if (!cards.length) return <div className={styles.status}>No cards found. Add items to the <strong>ContentCards</strong> list.</div>;

  const handleClick = (item: ICardItem): void => {
    if (!item.Url) return;
    if (item.OpenInNewTab) {
      window.open(item.Url, '_blank');
    } else {
      window.location.href = item.Url;
    }
  };

  return (
    <div className={styles.wrapper} dir="rtl">
      <div className={styles.header}>
        <span className={styles.handwriting}>מרגישים בבית</span>
        <h2 className={styles.sectionTitle}>כל השירותים בקליק</h2>
        <div className={styles.underline} />
      </div>

      <div className={styles.grid}>
        {cards.map(item => (
          <div
            key={item.Id}
            className={`${styles.card} ${item.IsHighlighted ? styles.highlighted : ''} ${item.IsLarge ? styles.large : ''}`}
            onClick={() => handleClick(item)}
            role="button"
            tabIndex={0}
            onKeyDown={e => e.key === 'Enter' && handleClick(item)}
          >
            {item.IsLarge ? (
              <div className={styles.cardContent}>
                {item.IconUrl && (
                  <img src={item.IconUrl} alt={item.Title} className={styles.icon} />
                )}
                <span className={styles.cardTitle}>{item.Title}</span>
                {item.SubTitle && (
                  <span className={styles.cardSubTitle}>{item.SubTitle}</span>
                )}
              </div>
            ) : (
              <>
                {item.IconUrl && (
                  <img src={item.IconUrl} alt={item.Title} className={styles.icon} />
                )}
                <div className={styles.cardContent}>
                  <span className={styles.cardTitle}>{item.Title}</span>
                  {item.SubTitle && (
                    <span className={styles.cardSubTitle}>{item.SubTitle}</span>
                  )}
                </div>
              </>
            )}
          </div>
        ))}
      </div>
    </div>
  );
};

export default Cards;
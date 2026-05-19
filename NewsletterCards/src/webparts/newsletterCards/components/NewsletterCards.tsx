import * as React from 'react';
import styles from './NewsletterCards.module.scss';
import type { INewsletterCardsProps } from './INewsletterCardsProps';
import {
  INewsletterItem,
  NewsletterCardsService
} from '../services/NewsletterCardsService';
import arrowImage from '../assets/arrow.png';

const NewsletterCards: React.FC<INewsletterCardsProps> = (props) => {
  const [items, setItems] = React.useState<INewsletterItem[]>([]);
  const [visibleCount, setVisibleCount] = React.useState<number>(2);
  const [isLoading, setIsLoading] = React.useState<boolean>(true);
  const [errorMessage, setErrorMessage] = React.useState<string>('');

  React.useEffect(() => {
    const loadItems = async (): Promise<void> => {
      try {
        const service = new NewsletterCardsService(props.context);
        const loadedItems = await service.getItems();

        setItems(loadedItems);
      } catch (error) {
        console.error(error);
        setErrorMessage('אירעה שגיאה בטעינת הנתונים מהרשימה.');
      } finally {
        setIsLoading(false);
      }
    };

    void loadItems();
  }, [props.context]);

  if (isLoading) {
    return (
      <section className={styles.newsletterWrapper} dir="rtl">
        <div className={styles.inner}>
          <div className={styles.message}>טוען נתונים...</div>
        </div>
      </section>
    );
  }

  if (errorMessage) {
    return (
      <section className={styles.newsletterWrapper} dir="rtl">
        <div className={styles.inner}>
          <div className={styles.errorMessage}>{errorMessage}</div>
        </div>
      </section>
    );
  }

  return (
    <section className={styles.newsletterWrapper} dir="rtl">
      <div className={styles.inner}>
        <h2 className={styles.title}>ניוזלטר ארגוני | כמה מילים...</h2>

        <div className={styles.cardsList}>
          {items.slice(0, visibleCount).map((item) => (
            <article className={styles.card} key={item.id}>
              <div className={styles.imageArea}>
                {item.imageUrl && <img src={item.imageUrl} alt={item.title} />}
              </div>

              <div className={styles.dateArea}>
                <div className={styles.calendarIcon} aria-hidden="true">
                  <span />
                </div>

                <div className={styles.dateText}>
                  <div>{item.hebrewDateLine1}</div>
                  <div>{item.hebrewDateLine2}</div>
                </div>

                <div className={styles.smallUnderline} />
              </div>

              <div className={styles.contentArea}>
                <h3>{item.title}</h3>
                <p>
                  {item.descriptionLine1}
                  <br />
                  {item.descriptionLine2}
                </p>
              </div>

              <a className={styles.viewLink} href={item.linkUrl}>
               <span className={styles.arrowCircle} aria-hidden="true">
                <img
                  className={styles.arrowImage}
                  src={arrowImage}
                  alt=""
                />
              </span>
                <span>לצפייה</span>
              </a>
            </article>
          ))}
        </div>

        {visibleCount < items.length && (
          <button
            className={styles.loadMoreButton}
            type="button"
            onClick={() => setVisibleCount((prev) => prev + 2)}
          >
            טען עוד
          </button>
        )}
      </div>
    </section>
  );
};

export default NewsletterCards;
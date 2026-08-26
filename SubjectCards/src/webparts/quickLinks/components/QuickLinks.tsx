import * as React from 'react';
import styles from './QuickLinks.module.scss';
import type { IQuickLinksProps } from './IQuickLinksProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { IQuickLinksItem, ListService } from '../services/ListService';

const QuickLinks: React.FC<IQuickLinksProps> = ({ listService }) => {
  const [quickLinks, setQuickLinks] = React.useState<IQuickLinksItem[]>([]);
  const [loading, setLoading] = React.useState(true);
  const [error, setError] = React.useState<string | null>(null);

  React.useEffect(() => {
    listService.getQuickLinks()
      .then(setQuickLinks)
      .catch(() => setError('Failed to load QuickLinks.'))
      .finally(() => setLoading(false));
  }, []);

  if (loading) return <div className={styles.status}>Loading Quick Links...</div>;
  if (error) return <div className={styles.status}>{error}</div>;
  if (!QuickLinks.length) return <div className={styles.status}>No Quick Links found. Add items to the <strong>QuickLinks</strong> list.</div>;

  const handleClick = (item: IQuickLinksItem): void => {
    const url = `${item.Url}`;
    if (item.OpenInNewTab) {
      window.open(url, '_blank');
    }
    else
      window.location.href = url;
  };

  return (
    <div>
      <div className={styles.container}>
        <h2 className={styles.title}>לינקים מהירים</h2>
        <div className={styles.list}>
          {quickLinks.map(item => (
            <div
              key={item.Id}
              className={styles.row}
              onClick={() => handleClick(item)}
              role="button"
              tabIndex={0}
              onKeyDown={e => e.key === 'Enter' && handleClick(item)}
            >

              
              {/* Title in the right */}
              <span className={styles.label}>{item.Title}</span>
              {/* Icon on the left */}
              <i className="fa-solid fa-arrow-up-right-from-square externalIcon"></i>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
};

export default QuickLinks;
 import * as React from 'react';
import { ISubjectCardsProps } from './ISubjectCardsProps';
import { IEmployee } from '../services/ListService';
import styles from './SubjectCards.module.scss';

const SortArrow: React.FC<{
  field: keyof IEmployee;
  activeField: keyof IEmployee | null;
  asc: boolean;
}> = ({ field, activeField, asc }) => {
  const isActive = activeField === field;
  const rotation = isActive && asc ? 180 : 0; // הופך כיוון לפי עולה/יורד

  return (
    <img
      src="https://ezermizionil.sharepoint.com/sites/portal/SiteAssets/HappyMoments/pic.png"
      alt=""
      className={styles.sortIcon}
      style={{ transform: `rotate(${rotation}deg)` }}
    />
  );
};

const EmployeeDirectory: React.FC<ISubjectCardsProps> = ({ listService }) => {
  const [allEmployees, setAllEmployees] = React.useState<IEmployee[]>([]);
  const [filtered, setFiltered] = React.useState<IEmployee[]>([]);
  const [loading, setLoading] = React.useState(true);
  const [error, setError] = React.useState<string | null>(null);
  const [search, setSearch] = React.useState('');
  const [sortField, setSortField] = React.useState<keyof IEmployee | null>(null);
  const [sortAsc, setSortAsc] = React.useState(true);

  React.useEffect(() => {
    listService.getEmployees()
      .then(data => {
        setAllEmployees(data);
        setFiltered(data);
      })
      .catch(() => setError('Failed to load employees.'))
      .finally(() => setLoading(false));
  }, []);

  const handleSearch = (value: string): void => {
    setSearch(value);
    if (!value) {
      setFiltered(allEmployees);
      return;
    }
    const lower = value.toLowerCase().trim();
    const parts = lower.split(' ').filter(Boolean);

    setFiltered(allEmployees.filter(e => {
      const fullName = `${(e.Title || '').toLowerCase()} ${(e.LastName || '').toLowerCase()}`;
      const fullNameReverse = `${(e.LastName || '').toLowerCase()} ${(e.Title || '').toLowerCase()}`;

      if (parts.length > 1) {
        return fullName.includes(lower) || fullNameReverse.includes(lower);
      }

      return (
        (e.Title || '').toLowerCase().includes(lower) ||
        (e.LastName || '').toLowerCase().includes(lower) ||
        (e.Email || '').toLowerCase().includes(lower) ||
        (e.Extension || '').toLowerCase().includes(lower) ||
        (e.Department || '').toLowerCase().includes(lower) ||
        (e.Branch || '').toLowerCase().includes(lower)
      );
    }));
  };

  const handleSort = (field: keyof IEmployee): void => {
    const asc = sortField === field ? !sortAsc : true;
    setSortField(field);
    setSortAsc(asc);
    const sorted = [...filtered].sort((a, b) => {
      const av = (a[field] || '') as string;
      const bv = (b[field] || '') as string;
      return asc ? av.localeCompare(bv, 'he') : bv.localeCompare(av, 'he');
    });
    setFiltered(sorted);
  };

  if (loading) return <div className={styles.status}>טוען...</div>;
  if (error) return <div className={styles.status}>{error}</div>;

  return (
    <div className={styles.wrapper} dir="rtl">
      <div className={styles.countBar}>
        <span className={styles.count}>
          {search
            ? `(מציג ${filtered.length} מתוך ${allEmployees.length} אנשי קשר)`
            : `(מציג ${allEmployees.length} מתוך ${allEmployees.length} אנשי קשר)`
          }
        </span>
      </div>
      <div className={styles.topBar}>
        <div className={styles.searchBox}>
          <span className={styles.searchIcon}>
            <svg width="16" height="16" viewBox="0 0 16 16" fill="none">
              <circle cx="6.5" cy="6.5" r="5.5" stroke="#36080A" strokeWidth="1.5"/>
              <line x1="10.5" y1="10.5" x2="14.5" y2="14.5" stroke="#36080A" strokeWidth="1.5" strokeLinecap="round"/>
            </svg>
          </span>
          <input
            type="text"
            placeholder="חיפוש חופשי"
            value={search}
            onChange={e => handleSearch(e.target.value)}
            className={styles.searchInput}
          />
          {search && (
            <button onClick={() => handleSearch('')} className={styles.clearBtn}>
               ניקוי חיפוש
            </button>
          )}
        </div>
        {search && (
          <button onClick={() => handleSearch('')} className={styles.clearBtnMobile}>
            ✕
          </button>
        )}
      </div>

      <div className={styles.desktopView}>
        <table className={styles.table}>
          <thead>
            <tr>
              <th><span className={styles.thBox}>שם פרטי</span></th>
              <th><span className={styles.thBox}>שם משפחה</span></th>
              <th><span className={styles.thBox}>שלוחה</span></th>
              <th><span className={styles.thBox}>אימייל</span></th>
              <th onClick={() => handleSort('Department')} className={styles.sortable}>
                <span className={styles.thBox}>
                  <span className={styles.thContent}>
                    <span>מחלקה</span>
                    <SortArrow field="Department" activeField={sortField} asc={sortAsc} />
                  </span>
                </span>
              </th>
              <th onClick={() => handleSort('Branch')} className={styles.sortable}>
                <span className={styles.thBox}>
                  <span className={styles.thContent}>
                    <span>סניף</span>
                    <SortArrow field="Branch" activeField={sortField} asc={sortAsc} />
                  </span>
                </span>
              </th>
            </tr>
          </thead>
          <tbody>
            {filtered.map(emp => (
              <tr key={emp.Id}>
                <td>{emp.Title}</td>
                <td>{emp.LastName}</td>
                <td>{emp.Extension}</td>
                <td>{emp.Email}</td>
                <td>{emp.Department}</td>
                <td>{emp.Branch}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <div className={styles.mobileView}>
        {filtered.map(emp => (
          <div key={emp.Id} className={styles.card}>
            <div className={styles.cardName}>{emp.Title} {emp.LastName}</div>
            <div className={styles.cardRow}><span>שלוחה:</span> {emp.Extension}</div>
            <div className={styles.cardRow}><span>אימייל:</span> {emp.Email}</div>
            <div className={styles.cardRow}><span>מחלקה:</span> {emp.Department}</div>
            <div className={styles.cardRow}><span>סניף:</span> {emp.Branch}</div>
          </div>
        ))}
      </div>
    </div>
  );
};

export default EmployeeDirectory;

 import * as React from 'react';
import styles from './Dashboard.module.scss';
import WidgetOne from './WidgetOne/WidgetOne';
import WidgetTwo from './WidgetTwo/WidgetTwo';
import { IDashboardProps } from './IDashboardProps';

const Dashboard: React.FC<IDashboardProps> = ({ listService, happyMomentsService }) => {
  return (
    <div className={styles.wrapper}>
      <div className={styles.one}>
        <WidgetOne listService={listService} />
      </div>
      <div className={styles.two}>
        <WidgetTwo listService={happyMomentsService} />
      </div>
    </div>
  );
};

export default Dashboard;
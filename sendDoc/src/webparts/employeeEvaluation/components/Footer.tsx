import * as React from 'react';
import { PrimaryButton } from '@fluentui/react';

interface IFooterProps {
  onSubmit: () => void;
  busy: boolean;
}

const Footer: React.FC<IFooterProps> = ({ onSubmit, busy }) => {
  return (
    <PrimaryButton
      text={busy ? 'מעדכן...' : 'התחלת תהליך הערכה'}
      onClick={onSubmit}
      disabled={busy}
    />
  );
};

export default Footer;
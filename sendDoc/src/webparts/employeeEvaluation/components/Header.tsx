import * as React from 'react';
import { Stack, Label } from '@fluentui/react';

interface IHeaderProps {
  title: string;
  subtitle?: string;
}

const Header: React.FC<IHeaderProps> = ({ title, subtitle }) => {
  return (
    <Stack horizontalAlign="center" styles={{ root: { marginBottom: 20 } }}>
      <Label styles={{ root: { fontSize: 24, fontWeight: 'bold' } }}>{title}</Label>
      {subtitle && <Label styles={{ root: { fontSize: 16, color: '#666' } }}>{subtitle}</Label>}
    </Stack>
  );
};

export default Header;
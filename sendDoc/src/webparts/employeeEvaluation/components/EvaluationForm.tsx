import * as React from 'react';
import {
  Checkbox,
  Dropdown,
  IDropdownOption,
} from '@fluentui/react';
import { PeoplePicker, PrincipalType } from '@pnp/spfx-controls-react/lib/PeoplePicker';
import { IUser } from './IEmployeeEvaluationProps';

interface IEvaluationFormProps {
  users: IUser[];
  rowSelection: Record<string, boolean>;
  setRowSelection: React.Dispatch<React.SetStateAction<Record<string, boolean>>>;
  peoplePickerContext: any;
  WORK_TYPE_OPTIONS: IDropdownOption[];
  setUserWorkType: React.Dispatch<React.SetStateAction<Record<string, string>>>;
}

const EvaluationForm: React.FC<IEvaluationFormProps> = ({
  users,
  rowSelection,
  setRowSelection,
  peoplePickerContext,
  WORK_TYPE_OPTIONS,
  setUserWorkType,
}) => {
  return (
    <div>
      {users.map((u) => (
        <div
          key={u.id}
          style={{
            display: 'grid',
            gridTemplateColumns: '40px 70px 70px',
            gap: 12,
            alignItems: 'center',
            padding: '2px 5px',
            borderBottom: '1px solid #e5e7eb',
          }}
        >
          {/* Select checkbox */}
          <Checkbox
            checked={!!rowSelection[u.id]}
            onChange={(_, checked) =>
              setRowSelection((prev) => ({ ...prev, [u.id]: !!checked }))
            }
          />

          {/* Email - PeoplePicker */}
          <PeoplePicker
            context={peoplePickerContext}
            webAbsoluteUrl={peoplePickerContext.pageContext.web.absoluteUrl}
            personSelectionLimit={1}
            principalTypes={[PrincipalType.User]}
            ensureUser={true}
            onChange={() => {}}
            showHiddenInUI={false}
          />

          {/* Work type */}
          <Dropdown
            options={WORK_TYPE_OPTIONS}
            onChange={(_, opt) => {
              if (opt) setUserWorkType((prev) => ({ ...prev, [u.id]: String(opt.key) }));
            }}
            styles={{ root: { minWidth: 120 } }}
          />
        </div>
      ))}
    </div>
  );
};

export default EvaluationForm;
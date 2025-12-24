import * as React from 'react';
import { Stack, Label, Checkbox, IconButton, Dropdown, IDropdownOption } from '@fluentui/react';
import { PeoplePicker, PrincipalType } from '@pnp/spfx-controls-react/lib/PeoplePicker';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IUser } from './IEmployeeEvaluationProps';

interface IManagerSelection {
  login?: string;
  displayName?: string;
}

interface IEvaluationListProps {
  selectedUsers: IUser[];
  onToggleSelectAllRows: (_: any, checked?: boolean) => void;
  rowSelection: Record<string, boolean>;
  setRowSelection: React.Dispatch<React.SetStateAction<Record<string, boolean>>>;
  // SharePoint data
  userEmployeeType: Record<string, string>;
  setUserEmployeeType: React.Dispatch<React.SetStateAction<Record<string, string>>>;
  userDepartment: Record<string, string>;
  setUserDepartment: React.Dispatch<React.SetStateAction<Record<string, string>>>;
  userSubDepartment: Record<string, string>;
  setUserSubDepartment: React.Dispatch<React.SetStateAction<Record<string, string>>>;
  selectedManagers: Record<string, {
    direct?: IManagerSelection | null;
    indirect?: IManagerSelection | null;
    operation?: IManagerSelection | null;
  }>;
  setSelectedManagers: React.Dispatch<React.SetStateAction<Record<string, {
    direct?: IManagerSelection | null;
    indirect?: IManagerSelection | null;
    operation?: IManagerSelection | null;
  }>>>;
  context: WebPartContext;
  departmentOptions: IDropdownOption[];
  subDepartmentOptions: IDropdownOption[];
  getSubDepartmentOptions: (dept: string) => IDropdownOption[];
  onSaveUser: (userId: string) => void;
  userActive: Record<string, boolean>;
  setUserActive: React.Dispatch<React.SetStateAction<Record<string, boolean>>>;

}

interface IEditingState {
  [userId: string]: boolean;
}

const EMPLOYEE_TYPE_OPTIONS: IDropdownOption[] = [
  { key: 'רגיל', text: 'רגיל' },
  { key: 'שעתי', text: 'שעתי' },
  { key: 'מנהל', text: 'מנהל' }
];

const EvaluationList: React.FC<IEvaluationListProps> = ({
  selectedUsers,
  onToggleSelectAllRows,
  rowSelection,
  setRowSelection,
  userEmployeeType,
  setUserEmployeeType,
  userDepartment,
  setUserDepartment,
  userSubDepartment,
  setUserSubDepartment,
  selectedManagers,
  setSelectedManagers,
  context,
  departmentOptions,
  subDepartmentOptions,
  getSubDepartmentOptions,
  onSaveUser,
  userActive,
  setUserActive,
}) => {
  const [editingIds, setEditingIds] = React.useState<IEditingState>({});

  const toggleEditMode = (userId: string) => {
    setEditingIds((prev) => ({
      ...prev,
      [userId]: !prev[userId],
    }));
  };

  const hasOwn = (o: object, k: string) =>
  Object.prototype.hasOwnProperty.call(o, k);


  return (
    <Stack tokens={{ childrenGap: 8 }}>

      {selectedUsers.length > 0 && (
        <Stack tokens={{ childrenGap: 6 }}>
          <Label>סה"כ {selectedUsers.length} עובדים:</Label>

          {/* בר עליון: בחר הכל + שיוך מרוכז */}
          <Stack tokens={{ childrenGap: 12 }}>
            <Checkbox
              label="בחר.י / בטל.י בחירת כל העובדים בטבלה"
              onChange={onToggleSelectAllRows}
            />
          </Stack>

          {/* כותרות הטבלה */}
          <div
            style={{
              display: 'grid',
              gridTemplateColumns: '40px 1fr 1fr 1fr 1fr 1fr 1fr 1fr 80px 40px 40px',
              gap: 12,
              alignItems: 'center',
              padding: '8px 5px',
              borderBottom: '2px solid #0078d4',
              backgroundColor: '#f3f2f1',
              fontWeight: 600,
            }}
          >
            <div></div>
            <Label style={{ fontWeight: 600 }}>שם עובד</Label>
            <Label style={{ fontWeight: 600 }}>סוג עובד</Label>
            <Label style={{ fontWeight: 600 }}>מחלקה</Label>
            <Label style={{ fontWeight: 600 }}>תת-מחלקה</Label>
            <Label style={{ fontWeight: 600 }}>מנהל ישיר</Label>
            <Label style={{ fontWeight: 600 }}>מנהל עקיף</Label>
            <Label style={{ fontWeight: 600 }}>מנהל מקצועי</Label>
            <Label style={{ fontWeight: 600 }}>פעיל</Label>
            <div></div>
            <div></div>
          </div>

          {/* רשימת עובדים */}
          {selectedUsers.map((u) => {
            const isEditing = editingIds[u.id];
            //const empType = userEmployeeType[u.id] || u.employeeType || '';
            //const dept = userDepartment[u.id] || u.department || '';
            //const subDept = userSubDepartment[u.id] || u.subDepartment || '';
            const empType = hasOwn(userEmployeeType, u.id) ? userEmployeeType[u.id] : (u.employeeType || '');
            const dept = hasOwn(userDepartment, u.id) ? userDepartment[u.id] : (u.department || '');
            const subDept = hasOwn(userSubDepartment, u.id) ? userSubDepartment[u.id] : (u.subDepartment || '');


            const rowSubDeptOptions = [
              { key: '', text: 'בחר תת-מחלקה' },
                ...getSubDepartmentOptions(dept).map(o => ({ ...o, selected: false })),
            ];

            console.log("😎rowSubDeptOptions ", rowSubDeptOptions);
            const managers = selectedManagers[u.id] || {};
            
            return (
              <div
                key={u.id}
                style={{
                  display: 'grid',
                  gridTemplateColumns: '40px 1fr 1fr 1fr 1fr 1fr 1fr 1fr 80px 40px 40px',
                  gap: 12,
                  alignItems: 'center',
                  padding: '8px 5px',
                  borderBottom: '1px solid #e5e7eb',
                  backgroundColor: isEditing ? '#fff4ce' : 'transparent',
                }}
              >
                <Checkbox
                  checked={!!rowSelection[u.id]}
                  onChange={(_, checked) =>
                    setRowSelection((prev) => ({ ...prev, [u.id]: !!checked }))
                  }
                />
                
                {/* שם עובד */}
                <Label>{u.displayName || u.userPrincipalName || '(ללא שם)'}</Label>
                
                {/* סוג עובד */}
                {isEditing ? (
                  <Dropdown
                    selectedKey={empType}
                    options={EMPLOYEE_TYPE_OPTIONS}
                    onChange={(_, option) => {
                      if (option) {
                        setUserEmployeeType((prev) => ({ ...prev, [u.id]: option.key as string }));
                      }
                    }}
                    styles={{ root: { minWidth: 100 } }}
                  />
                ) : (
                  <Label>{empType || '-'}</Label>
                )}
                
                {/* מחלקה */}
                {isEditing ? (
                  <Dropdown
                    selectedKey={dept}
                    options={departmentOptions}
                    onChange={(_, option) => {
                      if (option) {
                        setUserDepartment((prev) => ({ ...prev, [u.id]: option.key as string }));
                        setUserSubDepartment((prev) => ({ ...prev, [u.id]: '' }));

                      }
                    }}
                    styles={{ root: { minWidth: 120 } }}
                  />
                ) : (
                  <Label>{dept || '-'}</Label>
                )}
                
                {/* תת-מחלקה */}
                {isEditing ? (
                  <Dropdown
                    key={`${u.id}-${dept}`}
                    selectedKey={subDept? subDept: ''}
                    options={rowSubDeptOptions}
                    disabled={!dept}
                    onChange={(_, option) => {
                      if (option) {
                        setUserSubDepartment((prev) => ({ ...prev, [u.id]: option.key as string }));
                      }
                    }}
                    styles={{ root: { minWidth: 120 } }}
                  />
                ) : (
                  <Label>{subDept || '-'}</Label>
                )}
                
                {/* מנהל ישיר */}
                {isEditing ? (
                  <div style={{ minWidth: 150 }}>
                    <PeoplePicker
                      context={context as any}
                      webAbsoluteUrl={context.pageContext.web.absoluteUrl}
                      personSelectionLimit={1}
                      showtooltip={true}
                      required={false}
                      principalTypes={[PrincipalType.User]}
                      ensureUser={true}
                      resolveDelay={300}
                      defaultSelectedUsers={managers.direct?.login ? [managers.direct.login] : []}
                      onChange={(items) => {
                        const person = items && items.length > 0 ? items[0] : null;
                        setSelectedManagers((prev) => ({
                          ...prev,
                          [u.id]: {
                            ...prev[u.id],
                            direct: person ? { login: person.secondaryText, displayName: person.text } : null,
                          },
                        }));
                      }}
                    />
                  </div>
                ) : (
                  <Label>{managers.direct?.displayName || '-'}</Label>
                )}
                
                {/* מנהל עקיף */}
                {isEditing ? (
                  <div style={{ minWidth: 150 }}>
                    <PeoplePicker
                      context={context as any}
                      webAbsoluteUrl={context.pageContext.web.absoluteUrl}
                      personSelectionLimit={1}
                      showtooltip={true}
                      required={false}
                      principalTypes={[PrincipalType.User]}
                      ensureUser={true}
                      resolveDelay={300}
                      defaultSelectedUsers={managers.indirect?.login ? [managers.indirect.login] : []}
                      onChange={(items) => {
                        const person = items && items.length > 0 ? items[0] : null;
                        setSelectedManagers((prev) => ({
                          ...prev,
                          [u.id]: {
                            ...prev[u.id],
                            indirect: person ? { login: person.secondaryText, displayName: person.text } : null,
                          },
                        }));
                      }}
                    />
                  </div>
                ) : (
                  <Label>{managers.indirect?.displayName || '-'}</Label>
                )
                }


                {/* מנהל מקצועי */}
                {isEditing ? (
                  <div style={{ minWidth: 150 }}>
                    <PeoplePicker
                      context={context as any}
                      webAbsoluteUrl={context.pageContext.web.absoluteUrl}
                      personSelectionLimit={1}
                      showtooltip={true}
                      required={false}
                      principalTypes={[PrincipalType.User]}
                      ensureUser={true}
                      resolveDelay={300}
                      defaultSelectedUsers={managers.operation?.login ? [managers.operation.login] : []}
                      onChange={(items) => {
                        const person = items && items.length > 0 ? items[0] : null;
                        setSelectedManagers((prev) => ({
                          ...prev,
                          [u.id]: {
                            ...prev[u.id],
                            operation: person ? { login: person.secondaryText, displayName: person.text } : null,
                          },
                        }));
                      }}
                    />
                  </div>
                ) : (
                  <Label>{managers.operation?.displayName || '-'}</Label>
                )}

                {/* פעיל */}
                {isEditing ? (
                  <Checkbox
                    checked={!!userActive[u.id]}
                    onChange={(_, checked) =>
                      setUserActive((prev) => ({ ...prev, [u.id]: !!checked }))
                    }
                  />
                ) : (
                  <Checkbox checked={!!userActive[u.id]} disabled />
                )}


                
                <IconButton
                  iconProps={{ iconName: isEditing ? 'Cancel' : 'Edit' }}
                  title={isEditing ? 'ביטול' : 'עריכה'}
                  ariaLabel={isEditing ? 'ביטול' : 'עריכה'}
                  onClick={() => toggleEditMode(u.id)}
                  styles={{
                    root: { color: isEditing ? '#d13438' : '#0078d4' },
                  }}
                />
                <IconButton
                  iconProps={{ iconName: 'Save' }}
                  title="שמירה"
                  ariaLabel="שמירה"
                  disabled={!isEditing}
                  onClick={() => {
                    onSaveUser(u.id);
                    toggleEditMode(u.id);
                  }}
                  styles={{
                    root: { color: isEditing ? '#107C10' : '#CCCCCC' },
                  }}
                />
              </div>
            );
          })}
        </Stack>
      )}
    </Stack>
  );
};

export default EvaluationList;
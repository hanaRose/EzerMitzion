//emploee and direct users are users but by selecting a user it dosen't get a direct user and a group /

import * as React from 'react';
import {
  Stack, Label, Dropdown, IDropdownOption, PrimaryButton, MessageBar, MessageBarType, Checkbox, TextField
} from '@fluentui/react';
import { PeoplePicker, PrincipalType } from '@pnp/spfx-controls-react/lib/PeoplePicker';
import type { IPeoplePickerContext } from '@pnp/spfx-controls-react/lib/PeoplePicker';
import { IEmployeeEvaluationProps, IGroup, IUser } from './IEmployeeEvaluationProps';

// PnP module augmentations
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/fields';
import '@pnp/sp/items';
import '@pnp/sp/site-users/web';



const LIST_TITLE = 'adminEmployee';

const QUARTER_OPTIONS: IDropdownOption[] = [
  { key: 'Q1', text: 'Q1' },
  { key: 'Q2', text: 'Q2' },
  { key: 'Q3', text: 'Q3' },
  { key: 'Q4', text: 'Q4' }
];
// רשומת עובד כפי שהיא נשמרת ב־adminEmployee
type AdminEmployeeRow = {
  employeeType?: string;
  department?: string;
  subDepartment?: string;
  directManagerEmail?: string;
  directManagerTitle?: string;
  indirectManagerEmail?: string;
  indirectManagerTitle?: string;
  operationManagerEmail?: string;
  operationManagerTitle?: string;
};

const STATUS_CHOICES = [
  'ממתין לשליחה',
  'נשלח',
  'מולא ע"י העובד',
  'מולא על יד המנהל',
  'אושר',
  'נדחה',
  'נשלח לתיקון'
];

const WORK_TYPE_OPTIONS: IDropdownOption[] = [
  { key: 'רגיל', text: 'רגיל' },
  { key: 'שעתי', text: 'שעתי' },
  { key: 'מנהל', text: 'מנהל' }
];

const STATUS_OPTIONS: IDropdownOption[] = STATUS_CHOICES.map(s => ({ key: s, text: s }));


// ===== Helpers: normalize + token =====
const normalize = (s: string) =>
  (s || '')
    .toLowerCase()
    .normalize('NFKD')
    .replace(/[\u200E\u200F\u202A-\u202E]/g, '') // RTL marks
    .replace(/\([^)]*\)/g, ' ')                  // remove (dept) etc.
    .replace(/[^\p{L}\p{N}@.\s]+/gu, ' ')        // letters/digits/@/. and spaces
    .replace(/\s+/g, ' ')
    .trim();

const makeKey = (text: string, qName: string, qYear: string | number) =>
  `${normalize(text)}|${String(qName)}|${String(qYear)}`;




type UserMeta = {
  employeeType: string;
  managerDisplayName: string;
  managerLogin: string; // NEW: for ensureUser()
  indirectManagerEmail: string;
  operationManagerEmail: string;
  groupNamesForSelected: string[];
  department: string;
  subDepartment: string;
};

// מבנה מחלקה מהרשימה Departments
type DepartmentItem = {
  department: string;  // עמודה
  subDepartment: string;  // סוג
  address?: string;  // כתובת
};


const EmployeeEvaluation: React.FC<IEmployeeEvaluationProps> = (props) => {
  // עובדים שנבחרו ידנית מה-PeoplePicker
  const [manualUsers, setManualUsers] = React.useState<IUser[]>([]);

  // instance id to make console logs easy to find
  const instanceLogId = React.useRef<string>(`EZER-EE-${Date.now()}-${Math.random().toString(36).slice(2,8)}`);
  // expose id globally so you can query it in the console
  try { (window as any).__EZER_EVAL_ID = instanceLogId.current; } catch {}
  console.log(`>>> EZER-EVAL-CHECKPOINT: EmployeeEvaluation mounted. ID=${instanceLogId.current}`);
  // also log as error so it stands out in the console
  console.error(`*** EZER-EVAL-CHECKPOINT ERROR: mounted ID=${instanceLogId.current}`);

  // עובדים שבאים מקבוצות: gid -> רשימת עובדים
  const [groupUsersByGroup] = React.useState<Record<string, IUser[]>>({});

  const [userWorkType, setUserWorkType] = React.useState<Record<string, string>>({});
  const [userEmployeeName, setUserEmployeeName] = React.useState<Record<string, string>>({});
  const [userStatus, setUserStatus] = React.useState<Record<string, string>>({});
  // PeoplePicker selections for employee name and email
  const [selectedEmployeeName, setSelectedEmployeeName] = React.useState<Record<string, { login?: string; displayName?: string } | null>>({});
  const [selectedEmployeeEmail, setSelectedEmployeeEmail] = React.useState<Record<string, { login?: string; displayName?: string } | null>>({});
    // בחירת עובדים בטבלה לשיוך מרוכז
  const [rowSelection, setRowSelection] = React.useState<Record<string, boolean>>({});
  const [bulkWorkType, setBulkWorkType] = React.useState<string>('רגיל');

  // מחלקות ותת-מחלקות
  const [departmentsData, setDepartmentsData] = React.useState<DepartmentItem[]>([]);
  const [userDepartment, setUserDepartment] = React.useState<Record<string, string>>({});
  const [userSubDepartment, setUserSubDepartment] = React.useState<Record<string, string>>({});
  // per-user selected managers (direct / indirect / operation)
  const [selectedManagers, setSelectedManagers] = React.useState<Record<string, {
    direct?: { login?: string; displayName?: string } | null;
    indirect?: { login?: string; displayName?: string } | null;
    operation?: { login?: string; displayName?: string } | null;
  }>>({});

  const { sp, context } = props;
  const [groups] = React.useState<IGroup[]>([]);
  const [selectedGroupIds] = React.useState<string[]>([]);
  const [selectedUsers, setSelectedUsers] = React.useState<IUser[]>([]);
  const [busy, setBusy] = React.useState(false);
  const [msg, setMsg] = React.useState<{ type: MessageBarType; text: string } | null>(null);

  // “נשלח” לפי רבעון/שנה: טוקנים
   const [sentTokens, setSentTokens] = React.useState<Set<string>>(new Set());
  // const [ setGroupPreview] = React.useState<Record<string, GroupSentPreview>>({});
  // const [groupNewOnly, setGroupNewOnly] = React.useState<Record<string, boolean>>({});

  // רבעון/שנה ב-UI
  const [quarterName, setQuarterName] = React.useState<string>('Q1');
  const [quarterYear, setQuarterYear] = React.useState<string>(new Date().getFullYear().toString());

  // cache מטא למשתמש
  const userMetaCache = React.useRef<Map<string, UserMeta>>(new Map());

  const employeeNumberMapRef = React.useRef<Map<string, AdminEmployeeRow> | null>(null);

  // track which users we've logged from render to avoid noisy repeat logs
  const renderedLoggedUsersRef = React.useRef<Set<string>>(new Set());


  // שמות עמודות ה-User בפועל (אם קיימת התנגשות, נעבור לשמות גיבוי)
  const employeeUserFieldRef = React.useRef<string>('EmployeeUser');
  const managerUserFieldRef  = React.useRef<string>('DirectManager');
  const indirectManagerUserFieldRef = React.useRef<string>('IndirectManager');
  const operationManagerUserFieldRef = React.useRef<string>('OperationManager');

  // PeoplePicker context
  const peoplePickerContext: IPeoplePickerContext = {
    absoluteUrl: context.pageContext.web.absoluteUrl,
    spHttpClient: context.spHttpClient,
    msGraphClientFactory: context.msGraphClientFactory
  };

  const recomputeSelectedUsers = React.useCallback(() => {
    const byId = new Map<string, IUser>();

    // קודם עובדים ידניים
    manualUsers.forEach(u => {
      if (u?.id) byId.set(u.id, u);
    });

    // ואז כל העובדים מכל הקבוצות
    Object.values(groupUsersByGroup).forEach(arr => {
      arr.forEach(u => {
        if (u?.id && !byId.has(u.id)) {
          byId.set(u.id, u);
        }
      });
    });

    setSelectedUsers(Array.from(byId.values()));
  }, [manualUsers, groupUsersByGroup]);

  // helper: read a value from a per-user map trying both id and userPrincipalName
  const readUserMap = (map: Record<string, string>, u: IUser) => {
    const byId = u.id && map[u.id];
    const upn = (u.userPrincipalName || u.secondaryText || '').toLowerCase();
    const byUpn = upn && map[upn];
    return byId || byUpn || '';
  };

  React.useEffect(() => {
    recomputeSelectedUsers();
  }, [recomputeSelectedUsers]);

  // log when selectedUsers changes so we can see when rows become available
  React.useEffect(() => {
    try {
      console.error(`*** EZER-EVAL-CHECKPOINT ERROR: selectedUsers updated: ${selectedUsers.length} users ID=${instanceLogId.current}`);
    } catch {}
  }, [selectedUsers]);

  // --- יצירת אופציות למחלקות ---
  const departmentOptions = React.useMemo((): IDropdownOption[] => {
    const uniqueDepts = new Set<string>();
    departmentsData.forEach(d => {
      if (d.department) uniqueDepts.add(d.department);
    });
    return Array.from(uniqueDepts).map(dept => ({
      key: dept,
      text: dept
    }));
  }, [departmentsData]);

  React.useEffect(() => {
    (async () => {
      try {
        // רשימת המיפוי – לפי ה-GUID שנתת
        const dirList = sp.web.lists.getById('d0169395-ae9d-4173-a84a-dc3fd69d91c2');

        // חשוב: השמות כאן צריכים להתאים לשמות העמודות ברשימה!
        const items = await dirList.items
          .select('LinkTitle', 'field_6')
          .top(5000)(); // אפשר להגדיל אם צריך

        const m = new Map<string, AdminEmployeeRow>();

        for (const it of items) {
          const sam = (it.LinkTitle || '').toLowerCase().trim();
          const emp = (it.field_6 || '').toString().trim();
          if (!sam || !emp) continue;
          m.set(sam, emp);
        }

        console.log('📄 Loaded employeeNumber map from SP list:', m.size);
        employeeNumberMapRef.current = m;
      } catch (e) {
        console.warn('Failed to load employee numbers from SP list', e);
        employeeNumberMapRef.current = new Map();
      }
    })();
  }, [sp]);

  // helper: per-row manager change handler (PeoplePicker returns an array)
  const onRowManagerChange = (userId: string, role: 'direct' | 'indirect' | 'operation') => (items: any[]) => {
    const p = (items && items[0]) || null;
    const login = p?.secondaryText || p?.loginName || p?.id || '';
    const displayName = p?.text || p?.secondaryText || '';
    setSelectedManagers(prev => ({
      ...prev,
      [userId]: {
        ...(prev[userId] || {}),
        [role]: p ? { login, displayName } : null
      }
    }));
  };

  const onRowEmployeeDataChange = (userId: string, field: 'name' | 'email') => (items: any[]) => {
    const p = (items && items[0]) || null;
    const login = p?.secondaryText || p?.loginName || p?.id || '';
    const displayName = p?.text || p?.secondaryText || '';
    
    if (field === 'name') {
      setSelectedEmployeeName(prev => ({
        ...prev,
        [userId]: p ? { login, displayName } : null
      }));
    } else {
      setSelectedEmployeeEmail(prev => ({
        ...prev,
        [userId]: p ? { login, displayName } : null
      }));
    }
  };


  // --- טעינת מחלקות ותת-מחלקות ---
  React.useEffect(() => {
    (async () => {
      try {
        const deptList = sp.web.lists.getById('f1d888b2-f9a9-4b97-96f4-5216da5d50cc');

        const items = await deptList.items
          .select('Title', 'subDepartment')
          .top(5000)();

        const deptData: DepartmentItem[] = items.map((it: any) => ({
          department: it.Title || '',
          subDepartment: it.subDepartment || '',
          address: ''
        }));

        console.log('📊 Loaded departments:', deptData.length);
        console.log('📊 Unique departments:', new Set(deptData.map(d => d.department).filter(d => d)).size);
        console.log('📊 Sample data:', deptData.slice(0, 3));

        setDepartmentsData(deptData);
      } catch (e) {
        console.warn('Failed to load departments list', e);
        setDepartmentsData([]);
      }
    })();
  }, [sp]);

  // --- קבוצות מה-Graph ---
React.useEffect(() => {
  (async () => {
    try {
      // משתמשים ברשימה החדשה לפי שם – adminEmployee
      const dirList = sp.web.lists.getByTitle('adminEmployee');

      // בחר שדות רלוונטיים כולל user fields
      const items: any[] = await dirList.items
        .select(
         'Title',
    'employeeType',
    'WorkType',
    'EmployeeName',
    'Status',
    'department',
    'subDepartment',
    'employee/Title',
    'employee/EMail',
    'directManager/Title',
    'directManager/EMail',
    'indirectManager/Title',
    'indirectManager/EMail',
    'operationManager/Title',
    'operationManager/EMail'
        )
        .expand('employee', 'directManager', 'indirectManager', 'operationManager')
        .top(5000)();

      const map = new Map<string, AdminEmployeeRow>();
      const users: IUser[] = [];

      // Initialize state objects for all editable fields
      const workTypeMap: Record<string, string> = {};
      const employeeNameMap: Record<string, string> = {};
      const statusMap: Record<string, string> = {};
      const departmentMap: Record<string, string> = {};
      const subDepartmentMap: Record<string, string> = {};
      const managersMap: Record<string, {
        direct?: { login?: string; displayName?: string } | null;
        indirect?: { login?: string; displayName?: string } | null;
        operation?: { login?: string; displayName?: string } | null;
      }> = {};

      for (const it of items) {
        const sam = (it.Title || '').toLowerCase().trim();
        if (!sam) continue;

        map.set(sam, {
            employeeType: it.employeeType || '',
  department: it.department || '',
  subDepartment: it.subDepartment || '',

  directManagerEmail: it.directManager?.EMail || '',
  directManagerTitle: it.directManager?.Title || '',

  indirectManagerEmail: it.indirectManager?.EMail || '',
  indirectManagerTitle: it.indirectManager?.Title || '',

  operationManagerEmail: it.operationManager?.EMail || '',
  operationManagerTitle: it.operationManager?.Title || ''
        });

        // Build a user entry for the table. Prefer the expanded employee user if present.
        const email = it.employee?.EMail || '';
        const display = it.employee?.Title || it.Title || email || sam;
        const idKey = email || it.Title || sam;

        users.push({
          id: String(idKey),
          displayName: display,
          userPrincipalName: email.toLowerCase(),
          secondaryText: email
        });

        // Populate state maps with existing values from the list
        const userId = String(idKey);
        // Use WorkType or employeeType as fallback (some rows store the type in employeeType)
        if (it.WorkType || it.employeeType) workTypeMap[userId] = it.WorkType || it.employeeType;
        if (it.EmployeeName) employeeNameMap[userId] = it.EmployeeName;
        if (it.Status) statusMap[userId] = it.Status;
        if (it.department) departmentMap[userId] = it.department;
        if (it.subDepartment) subDepartmentMap[userId] = it.subDepartment;

        // Populate managers
        managersMap[userId] = {
          direct: it.directManager?.EMail ? {
            login: it.directManager.EMail,
            displayName: it.directManager.Title || it.directManager.EMail
          } : null,
          indirect: it.indirectManager?.EMail ? {
            login: it.indirectManager.EMail,
            displayName: it.indirectManager.Title || it.indirectManager.EMail
          } : null,
          operation: it.operationManager?.EMail ? {
            login: it.operationManager.EMail,
            displayName: it.operationManager.Title || it.operationManager.EMail
          } : null
        };
      }

      console.log('Loaded adminEmployee directory rows:', map.size);
      employeeNumberMapRef.current = map;

      // conspicuous checkpoint so user can find this load in console
      console.log(`>>> EZER-EVAL-CHECKPOINT: adminEmployee rows loaded: ${map.size} ID=${instanceLogId.current}`);

      // במקום בחירת ידנית — נטען את כל העובדים מהרשימה כאילו נבחרו
      setManualUsers(users);
      // conspicuous log so user can spot when manual users are set
      try {
        console.error(`*** EZER-EVAL-CHECKPOINT ERROR: setManualUsers called with ${users.length} users ID=${instanceLogId.current}`);
      } catch {}

      // log selected managers map size and a sample of keys
      try {
        console.error(`*** EZER-EVAL-CHECKPOINT ERROR: setting selectedManagers for ${Object.keys(managersMap).length} users ID=${instanceLogId.current}`, Object.keys(managersMap).slice(0,10));
      } catch {}

      // Set all the state with loaded values
      setUserWorkType(workTypeMap);
      setUserEmployeeName(employeeNameMap);
      setUserStatus(statusMap);
      setUserDepartment(departmentMap);
      setUserSubDepartment(subDepartmentMap);
      setSelectedManagers(managersMap);
    } catch (e) {
      console.warn('Failed to load employee directory from adminEmployee list', e);
      employeeNumberMapRef.current = new Map();
    }
  })();
}, [sp]);

  // --- טעינת “נשלח” מהרשימה (כולל רבעון/שנה) ---
  React.useEffect(() => {
    (async () => {
      try {
        const list = sp.web.lists.getByTitle(LIST_TITLE);
        const items = await list.items
          .select('Id','Title','EmployeeName','QuarterName','QuarterYear')
          .top(5000)();

        const tokens = new Set<string>();
        for (const it of items) {
          const qn = String(it.QuarterName ?? '');
          const qy = String(it.QuarterYear ?? '');
          if (it.Title)        tokens.add(makeKey(it.Title,        qn, qy));
          if (it.EmployeeName) tokens.add(makeKey(it.EmployeeName, qn, qy));
        }
        setSentTokens(tokens);
      } catch {
        setSentTokens(new Set());
      }
    })();
  }, [sp]);

  // --- PeoplePicker removed - employees are loaded automatically from adminEmployee list ---
const ensureUserField = async (
  list: any,
  preferredInternalName: string,
  fallbackInternalName: string,
  description: string
) => {
  // נסה להביא שדה קיים בשם המועדף
  try {
    const f = await list.fields
      .getByInternalNameOrTitle(preferredInternalName)
      .select('InternalName', 'TypeAsString')();

    if (f?.TypeAsString === 'User') {
      // יש שדה User בשם המועדף – להשתמש בו
      return f.InternalName; // מחזיר את ה-InternalName האמיתי!
    }
    // קיים אבל לא מטיפוס User – נשתמש בגיבוי
  } catch {
    // לא קיים – ננסה ליצור בשם המועדף
    try {
      const created = await list.fields.addUser(preferredInternalName, {
        Description: description,
        SelectionMode: 0 // Single user
      });
      return created.data?.InternalName || preferredInternalName;
    } catch {
      // ייתכן שנכשל מסיבה אחרת – נמשיך לייצר גיבוי
    }
  }

  // גיבוי: EmployeeUser / DirectManagerUser
  try {
    const f2 = await list.fields
      .getByInternalNameOrTitle(fallbackInternalName)
      .select('InternalName', 'TypeAsString')();

    if (f2?.TypeAsString === 'User') {
      return f2.InternalName; // מחזיר את ה-InternalName האמיתי!
    }
  } catch {
    // לא קיים – ניצור
  }

  const created2 = await list.fields.addUser(fallbackInternalName, {
    Description: description,
    SelectionMode: 0
  });

  return created2.data?.InternalName || fallbackInternalName;
};

  const ensureList = async () => {
      // בדיקה אם הרשימה קיימת, ואם לא – יצירה
      let listExists = true;
      try {
        await sp.web.lists.getByTitle(LIST_TITLE)();
      } catch {
        listExists = false;
      }

      if (!listExists) {
        await sp.web.lists.add(LIST_TITLE, 'Workers created by SPFx', 100, true);
      }

      const list = sp.web.lists.getByTitle(LIST_TITLE);

      // --- עזר קטן: הבטחת שדה לפי שם (InternalName או Title) ---

      const ensureTextField = async (nameOrTitle: string, opts?: any) => {
        try {
          await list.fields.getByInternalNameOrTitle(nameOrTitle)();
        } catch {
          await list.fields.addText(nameOrTitle, opts || {});
        }
      };

      const ensureChoiceField = async (nameOrTitle: string, opts: any) => {
        try {
          await list.fields.getByInternalNameOrTitle(nameOrTitle)();
        } catch {
          await list.fields.addChoice(nameOrTitle, opts);
        }
      };

      const ensureNumberField = async (nameOrTitle: string) => {
        try {
          await list.fields.getByInternalNameOrTitle(nameOrTitle)();
        } catch {
          await list.fields.addNumber(nameOrTitle);
        }
      };



      await ensureChoiceField('WorkType', {
        Choices: ['רגיל', 'שעתי', 'מנהל'],
        FillInChoice: false
      });

      // --- שדות טקסט/בחירה/מספר ---

      await ensureTextField('EmployeeName', {
        Description: 'שם העובד'
      });

      await ensureTextField('department', {
        Description: 'מחלקה',
        MaxLength: 255
      });

      await ensureTextField('subDepartment', {
        Description: 'תת-מחלקה',
        MaxLength: 255
      });

      await ensureChoiceField('employeeType', {
        Choices: ['רגיל', 'שעתי', 'מנהל'],
        FillInChoice: false
      });

      // אם כבר יצרת בעבר DirectManager כטקסט — לא נוגעים בו כאן; יהיה שדה User נפרד בהמשך

      await ensureChoiceField('QuarterName', {
        Choices: ['Q1', 'Q2', 'Q3', 'Q4'],
        FillInChoice: false
      });

      await ensureNumberField('QuarterYear');

      await ensureChoiceField('Status', {
        Choices: STATUS_CHOICES,
        FillInChoice: false
      });

      // ברירת מחדל ל-Status
      try {
        await list.fields
          .getByInternalNameOrTitle('Status')
          .update({ DefaultValue: 'ממתין לשליחה' });
      } catch {
        // לא קריטי אם נכשל
      };

      // --- הבטחת עמודות User אמיתיות לעובד ולמנהל ---
      // אם "Employee" או "DirectManager" קיימים בטיפוס שגוי — ניצור EmployeeUser / DirectManagerUser

      const employeeField = await ensureUserField(
        list,
        'employee',
        'Employee',
        'העובד הנבחר'
      );

      const managerField = await ensureUserField(
        list,
        'directManager',
        'DirectManager',
        'המנהל הישיר'
      );

      const indirectManagerField = await ensureUserField(
        list,
        'indirectManager',
        'IndirectManager',
        'המנהל העקיף'
      );

      const operationManagerField = await ensureUserField(
        list,
        'operationManager',
        'OperationManager',
        'מנהל התפעול'
      );

      employeeUserFieldRef.current = employeeField;
      managerUserFieldRef.current = managerField;
      indirectManagerUserFieldRef.current = indirectManagerField;
      operationManagerUserFieldRef.current = operationManagerField;

      console.log('Field names:', {
        employee: employeeField,
        manager: managerField,
        indirectManager: indirectManagerField,
        operationManager: operationManagerField
      });

      try {
        console.error(`*** EZER-EVAL-CHECKPOINT ERROR: ensured list user fields ID=${instanceLogId.current}`, {
          employeeField, managerField, indirectManagerField, operationManagerField
        });
      } catch {}

      // בדיקה: איזה שדות באמת קיימים?
      try {
        const allFields = await list.fields.filter('TypeAsString eq \'User\'').select('InternalName', 'Title', 'TypeAsString')();
        console.log('All User fields in list:', allFields);
      } catch (e) {
        console.warn('Could not fetch all fields', e);
      }
  };



  // --- מטא־דאטה אוטומטי למשתמש ---
const getUserMeta = async (user: IUser): Promise<UserMeta> => {
  const cacheKey = user.id || user.userPrincipalName;
  if (cacheKey && userMetaCache.current.has(cacheKey)) {
    return userMetaCache.current.get(cacheKey)!;
  }

  // ערכי ברירת מחדל אם אין התאמה ברשימה
  let employeeType = 'רגיל';
  let managerDisplayName = '';
  let managerLogin = '';
  let indirectManagerEmail = '';
  let operationManagerEmail = '';
  const groupNamesForSelected: string[] = []; // אין צורך בקבוצות כרגע
  let department = '';
  let subDepartment = '';

  try {
    if (employeeNumberMapRef.current) {
      const upn = (user.userPrincipalName || user.secondaryText || '').toLowerCase().trim();
      if (upn) {
        const sam = upn.split('@')[0]; // "user@domain" -> "user"
        const row = employeeNumberMapRef.current.get(sam);

        if (row) {
          employeeType           = row.employeeType || employeeType;
          department             = row.department || '';
          subDepartment          = row.subDepartment || '';
          managerDisplayName     = row.directManagerTitle || row.directManagerEmail || '';
          managerLogin           = row.directManagerEmail || '';
          indirectManagerEmail   = row.indirectManagerEmail || '';
          operationManagerEmail  = row.operationManagerEmail || '';
        }
      }
    }
  } catch (e) {
    console.warn('Failed to resolve meta from adminEmployee list for user', user, e);
  }

  const meta: UserMeta = {
    employeeType,
    managerDisplayName,
    managerLogin,
    indirectManagerEmail,
    operationManagerEmail,
    groupNamesForSelected, // נשאר ריק
    department,
    subDepartment
  };

  if (cacheKey) {
    userMetaCache.current.set(cacheKey, meta);
  }

  return meta;
};

  // --- הוספת/עדכון פריט (כפילות נחסמת לפי רבעון/שנה נוכחיים) ---
  const addWorkerItemIfMissing = async (user: IUser, source: string, groupId?: string) => {
    const list = sp.web.lists.getByTitle(LIST_TITLE);

    const upnRaw = (user.userPrincipalName || user.displayName || '');
    const upnEsc = upnRaw.replace(/'/g, "''");

    const qnEsc = quarterName.replace(/'/g, "''");
    const qyNum = parseInt(quarterYear, 10) || new Date().getFullYear();

    // בדיקת כפילות *באותו* רבעון/שנה
    const filter = `Title eq '${upnEsc}' and QuarterName eq '${qnEsc}' and QuarterYear eq ${qyNum}`;
    const existing = await list.items.filter(filter).top(1)();

    const meta = await getUserMeta(user);
    const groupNameString = meta.groupNamesForSelected.join(', ');

    // key used to index per-user maps (id or upn)
    const userKey = String(user.id || user.userPrincipalName || user.displayName || '').toLowerCase();


    const workType = readUserMap(userWorkType, user);
    const employeeName = (readUserMap(userEmployeeName, user) || user.displayName || user.userPrincipalName || '');
    const statusValue = (readUserMap(userStatus, user) || 'ממתין לשליחה');

    // הבטחת Site Users Ids לעובד ולמנהל
    const employeeLogin = user.userPrincipalName || user.displayName || '';


    const ensuredEmployee = await sp.web.ensureUser(employeeLogin);
    const employeeUserId = ensuredEmployee.Id;

    // Resolve managers: prefer user-selected managers (per-row) over meta-derived values
    let directManagerUserId: number | null = null;
    let indirectManagerUserId: number | null = null;
    let operationManagerUserId: number | null = null;

    const selManagers = userKey ? selectedManagers[userKey] : undefined;

    // direct
    if (selManagers?.direct?.login) {
      try {
        const ens = await sp.web.ensureUser(selManagers.direct.login);
        directManagerUserId = ens.Id;
      } catch (e) {
        console.warn('Failed to ensure selected direct manager user:', selManagers.direct.login, e);
        directManagerUserId = null;
      }
    } else if (meta.managerLogin) {
      try {
        const ensuredManager = await sp.web.ensureUser(meta.managerLogin);
        directManagerUserId = ensuredManager.Id;
      } catch {
        directManagerUserId = null;
      }
    }

    // indirect
    if (selManagers?.indirect?.login) {
      try {
        const ens = await sp.web.ensureUser(selManagers.indirect.login);
        indirectManagerUserId = ens.Id;
      } catch (e) {
        console.warn('Failed to ensure selected indirect manager user:', selManagers.indirect.login, e);
        indirectManagerUserId = null;
      }
    } else if (meta.indirectManagerEmail) {
      try {
        const ensuredIndirectManager = await sp.web.ensureUser(meta.indirectManagerEmail);
        indirectManagerUserId = ensuredIndirectManager.Id;
      } catch (e) {
        console.warn('Failed to ensure indirect manager user:', meta.indirectManagerEmail, e);
        indirectManagerUserId = null;
      }
    }

    // operation
    if (selManagers?.operation?.login) {
      try {
        const ens = await sp.web.ensureUser(selManagers.operation.login);
        operationManagerUserId = ens.Id;
      } catch (e) {
        console.warn('Failed to ensure selected operation manager user:', selManagers.operation.login, e);
        operationManagerUserId = null;
      }
    } else if (meta.operationManagerEmail) {
      try {
        const ensuredOperationManager = await sp.web.ensureUser(meta.operationManagerEmail);
        operationManagerUserId = ensuredOperationManager.Id;
      } catch (e) {
        console.warn('Failed to ensure operation manager user:', meta.operationManagerEmail, e);
        operationManagerUserId = null;
      }
    }

    // שמות השדות בפועל (ייתכן שהם EmployeeUser / DirectManagerUser)
    const employeeFieldName = employeeUserFieldRef.current;   // e.g. 'Employee' or 'EmployeeUser'

    // מחלקה ותת-מחלקה של העובד הספציפי
    const userDept = userKey ? userDepartment[userKey] : '';
    const userSubDept = userKey ? userSubDepartment[userKey] : '';

    const baseFields: any = {
      Title: upnRaw,
      EmployeeName: employeeName,
      employeeType: workType,
      QuarterName: quarterName,
      QuarterYear: qyNum,
      Status: statusValue,
      GroupName: groupNameString,
      WorkType: workType,
      department: userDept || meta.department || '',
      subDepartment: userSubDept || meta.subDepartment || ''
    };

    // הוסף User fields ל-baseFields (עם Id בסוף)
    if (employeeUserId) {
      baseFields[`${employeeFieldName}Id`] = employeeUserId;
    }
    if (directManagerUserId) {
      baseFields[`${managerUserFieldRef.current}Id`] = directManagerUserId;
    }
    if (indirectManagerUserId) {
      baseFields[`${indirectManagerUserFieldRef.current}Id`] = indirectManagerUserId;
    }
    if (operationManagerUserId) {
      baseFields[`${operationManagerUserFieldRef.current}Id`] = operationManagerUserId;
    }

    // Coerce known string fields to strings to avoid Edm.String conversion errors
    const stringFields = ['Title','EmployeeName','employeeType','QuarterName','Status','GroupName','WorkType','department','subDepartment','employeeId'];
    for (const key of stringFields) {
      if (Object.prototype.hasOwnProperty.call(baseFields, key)) {
        const v = baseFields[key];
        if (v === undefined || v === null) baseFields[key] = '';
        else if (typeof v !== 'string') baseFields[key] = String(v);
      }
    }

    console.debug('Adding item with all fields:', baseFields);

    if (existing.length === 0) {
      console.debug('Creating new item in list', LIST_TITLE);

      // יצירת הפריט עם כל השדות כולל User fields
      const addResult = await list.items.add(baseFields);
      const newItemId = addResult.data?.Id || addResult.Id;

      console.debug('Item created successfully with ID:', newItemId);
    } else {
      console.debug('Item already exists, updating instead. ID:', existing[0].Id);
      const id = existing[0].Id;
      const updateFields: any = {
        EmployeeName: employeeName,
        employeeType: workType,
        WorkType: workType,
        Status: statusValue,
        department: userDept || meta.department || existing[0].department || '',
        subDepartment: userSubDept || meta.subDepartment || existing[0].subDepartment || ''
      };

      // הוסף User fields (עם Id בסוף)
      if (employeeUserId) {
        updateFields[`${employeeFieldName}Id`] = employeeUserId;
      }
      if (directManagerUserId) {
        updateFields[`${managerUserFieldRef.current}Id`] = directManagerUserId;
      }
      if (indirectManagerUserId) {
        updateFields[`${indirectManagerUserFieldRef.current}Id`] = indirectManagerUserId;
      }
      if (operationManagerUserId) {
        updateFields[`${operationManagerUserFieldRef.current}Id`] = operationManagerUserId;
      }

      // Ensure update fields are strings where SharePoint expects strings
      const updateStringFields = ['EmployeeName','employeeType','WorkType','Status','department','subDepartment','employeeId'];
      for (const key of updateStringFields) {
        if (Object.prototype.hasOwnProperty.call(updateFields, key)) {
          const v = updateFields[key];
          if (v === undefined || v === null) updateFields[key] = '';
          else if (typeof v !== 'string') updateFields[key] = String(v);
        }
      }

      console.debug('Updating existing item with fields:', updateFields);
      await list.items.getById(id).update(updateFields);
      console.debug('Successfully updated item');
    }
  };

  // --- מעטפת שממשיכה גם כשיש שגיאה למשתמש בודד ---
  const tryAddWorker = async (user: IUser, source: string, groupId?: string) => {
    try {
      await addWorkerItemIfMissing(user, source, groupId);
      return { ok: true as const, user };
    } catch (e: any) {
      console.warn('Failed for user', user, e);
      return { ok: false as const, user, error: e };
    }
  };

  // --- שליחה ---
  const onSubmit = async () => {
    setMsg(null);
    setBusy(true);
    try {
      if (!/^\d{4}$/.test(quarterYear)) {
        setMsg({ type: MessageBarType.error, text: 'אנא הזיני שנת רבעון בת 4 ספרות (לדוגמה: 2025).' });
        setBusy(false);
        return;
      }

            // ✅ בדיקה: אין עובד ללא סוג עובד
      const usersWithoutType = selectedUsers.filter(u => !readUserMap(userWorkType, u));

      if (usersWithoutType.length > 0) {
        const names = usersWithoutType
          .slice(0, 10)
          .map(u => u.displayName || u.userPrincipalName || '(ללא שם)')
          .join(', ');

        const extra = usersWithoutType.length > 10
          ? ` ועוד ${usersWithoutType.length - 10} נוספים`
          : '';

        setMsg({
          type: MessageBarType.error,
          text: `העובד/ים הבא/ים לא שויכו לסוג עובד ולכן לא ניתן לשמור: ${names}${extra}. יש לשייך סוג עובד לכל העובדים לפני שמירה.`
        });
        setBusy(false);
        return;
      }


      await ensureList();

      const actuallySent: IUser[] = [];
      const failures: { user: IUser; error: any }[] = [];

      // 1) משתמשים נבחרים — מעדכן תמיד את כל הרשומות (יוצר חדשות או מעדכן קיימות)
      const manualById = new Map<string, IUser>();
      for (const u of manualUsers) {
        if (u?.id) manualById.set(u.id, u);
      }
      for (const u of Array.from(manualById.values())) {
        const r = await tryAddWorker(u, 'Selected', undefined);
        if (r.ok) actuallySent.push(u);
        else failures.push({ user: u, error: r.error });
      }

      // 2) קבוצות (מסונן לפי sentTokens לרבעון/שנה הנוכחיים)
      for (const gid of selectedGroupIds) {
        const g = groups.find(x => x.id === gid);
        const gName = g?.displayName ?? gid;
          const members: IUser[] = groupUsersByGroup[gid] || [];
          if (members.length === 0) {
            continue; // אין עובדים בקבוצה הזו כרגע
          }


       
        const membersToSend = true
          ? members.filter(m => {
              const k1 = makeKey(m.userPrincipalName || '', quarterName, quarterYear);
              const k2 = makeKey(m.displayName || '',       quarterName, quarterYear);
              return !(sentTokens.has(k1) || sentTokens.has(k2));
            })
          : members;

        for (const u of membersToSend) {
          const r = await tryAddWorker(u, `FromGroup:${gName}`, gid);
          if (r.ok) actuallySent.push(u);
          else failures.push({ user: u, error: r.error });
        }

        // await ensureGroupPreview(gid);
      }

      // עדכון sentTokens רק עבור מי שבאמת נשלח (ברבעון/שנה הנוכחיים)
      const newSent = new Set(sentTokens);
      for (const u of actuallySent) {
        if (u.userPrincipalName) newSent.add(makeKey(u.userPrincipalName, quarterName, quarterYear));
        if (u.displayName)       newSent.add(makeKey(u.displayName,       quarterName, quarterYear));
      }
      setSentTokens(newSent);

      // הודעת סיכום
      if (failures.length === 0) {
        setMsg({ type: MessageBarType.success, text: `עודכנו בהצלחה ${actuallySent.length} רשומות עובדים (נוצרו חדשות או עודכנו קיימות).` });
      } else {
        const names = failures
          .slice(0, 10)
          .map(f => f.user.displayName || f.user.userPrincipalName || '(ללא שם)')
          .join(', ');
        const extra = failures.length > 10 ? ` ועוד ${failures.length - 10} נוספים` : '';
        setMsg({
          type: MessageBarType.warning,
          text: `הפעולה הושלמה חלקית: ${actuallySent.length} עובדים עודכנו בהצלחה, אך ${failures.length} כשלו. בעיות: ${names}${extra}. ראי לוג בקונסול לפרטים.`
        });
      }
    } catch (e: any) {
      setMsg({ type: MessageBarType.error, text: `שגיאה בשליחה: ${e?.message || e}` });
    } finally {
      setBusy(false);
    }
  };

  // ====== PeoplePicker highlighting removed - no longer needed ======

    const onToggleSelectAllRows = (_: any, checked?: boolean) => {
    const next: Record<string, boolean> = {};
    if (checked) {
      selectedUsers.forEach(u => {
        if (u?.id) next[u.id] = true;
      });
    }
    setRowSelection(next);
  };


  const renderUserBadge = (u: IUser) => {
    // log per-user manager data once to help debugging
    try {
      if (u?.id && !renderedLoggedUsersRef.current.has(u.id)) {
        console.error(`*** EZER-EVAL-CHECKPOINT ERROR: renderUserBadge for user ${u.id} ID=${instanceLogId.current}`, {
          selectedManagers: selectedManagers?.[u.id] || null,
          userWorkType: userWorkType?.[u.id] || null,
          userDepartment: userDepartment?.[u.id] || null
        });
        renderedLoggedUsersRef.current.add(u.id);
      }
    } catch {}
    const already =
      sentTokens.has(makeKey(u.userPrincipalName || '', quarterName, quarterYear)) ||
      sentTokens.has(makeKey(u.displayName || '',       quarterName, quarterYear));

    const currentWorkType = readUserMap(userWorkType, u) || '';
    const currentDepartment = readUserMap(userDepartment, u) || '';
    const currentSubDepartment = readUserMap(userSubDepartment, u) || '';
    const currentStatus = readUserMap(userStatus, u) || 'ממתין לשליחה';

    const getUserSubDeptOptions = (): IDropdownOption[] => {
      if (!currentDepartment) return [];
      const uniqueSubDepts = new Set<string>();
      departmentsData
        .filter(d => d.department === currentDepartment)
        .forEach(d => { if (d.subDepartment) uniqueSubDepts.add(d.subDepartment); });
      return Array.from(uniqueSubDepts).map(s => ({ key: s, text: s }));
    };
    const userSubDeptOptions = getUserSubDeptOptions();

    return (
      <div
        key={u.id}
        style={{
          display: 'grid',
          gridTemplateColumns: '40px 70px 70px 70px 70px 70px 70px 70px 70px 70px',
          gap: 12,
          alignItems: 'center',
          padding: '2px 5px',
          borderBottom: '1px solid #e5e7eb',
          background: already ? '#ffeef4' : 'transparent'
        }}
      >
        {/* select checkbox */}
        <div style={{ display: 'flex', alignItems: 'center' }}>
          <Checkbox
            checked={!!rowSelection[u.id]}
            onChange={(_, checked) => setRowSelection(prev => ({ ...prev, [u.id]: !!checked }))}
          />
        </div>


        {/* email - PeoplePicker */}
        <div>
          <PeoplePicker
            context={peoplePickerContext}
            personSelectionLimit={1}
            principalTypes={[PrincipalType.User]}
            ensureUser={true}
            onChange={onRowEmployeeDataChange(u.id, 'email')}
            showHiddenInUI={false}
          />
        </div>

        {/* work type */}
        <div>
          <Dropdown
            options={WORK_TYPE_OPTIONS}
            selectedKey={currentWorkType ? currentWorkType : undefined}
            onChange={(_, opt) => {
              if (opt) setUserWorkType(prev => ({ ...prev, [u.id]: String(opt.key) }));
            }}
            styles={{ root: { minWidth: 120 } }}
          />
        </div>

        {/* employeeType removed (managed via WorkType) */}

        {/* department */}
        <div>
          <Dropdown
            options={departmentOptions}
            placeholder="בחר מחלקה"
            selectedKey={currentDepartment || undefined}
            onChange={(_, opt) => {
              setUserDepartment(prev => ({ ...prev, [u.id]: (opt?.key as string) || '' }));
              setUserSubDepartment(prev => ({ ...prev, [u.id]: '' }));
            }}
            styles={{ root: { minWidth: 140 } }}
          />
        </div>

        {/* sub-department */}
        <div>
          <Dropdown
            options={userSubDeptOptions}
            placeholder={currentDepartment ? 'בחר תת-מחלקה' : 'בחר מחלקה תחילה'}
            selectedKey={currentSubDepartment || undefined}
            disabled={!currentDepartment}
            onChange={(_, opt) => setUserSubDepartment(prev => ({ ...prev, [u.id]: (opt?.key as string) || '' }))}
            styles={{ root: { minWidth: 140 } }}
          />
        </div>

        {/* status */}
        <div>
          <Dropdown
            options={STATUS_OPTIONS}
            placeholder="בחר סטטוס"
            selectedKey={currentStatus || undefined}
            onChange={(_, opt) => setUserStatus(prev => ({ ...prev, [u.id]: (opt?.key as string) || '' }))}
            styles={{ root: { minWidth: 150 } }}
          />
        </div>

        {/* direct manager */}
        <div>
          <PeoplePicker
            context={peoplePickerContext}
            personSelectionLimit={1}
            principalTypes={[PrincipalType.User]}
            ensureUser={true}
            defaultSelectedUsers={selectedManagers?.[u.id]?.direct?.login ? [selectedManagers[u.id].direct!.login!] : []}
            onChange={onRowManagerChange(u.id, 'direct')}
            showHiddenInUI={false}
          />
        </div>

        {/* indirect manager */}
        <div>
          <PeoplePicker
            context={peoplePickerContext}
            personSelectionLimit={1}
            principalTypes={[PrincipalType.User]}
            ensureUser={true}
            defaultSelectedUsers={selectedManagers?.[u.id]?.indirect?.login ? [selectedManagers[u.id].indirect!.login!] : []}
            onChange={onRowManagerChange(u.id, 'indirect')}
            showHiddenInUI={false}
          />
        </div>

        {/* operation manager */}
        <div>
          <PeoplePicker
            context={peoplePickerContext}
            personSelectionLimit={1}
            principalTypes={[PrincipalType.User]}
            ensureUser={true}
            defaultSelectedUsers={selectedManagers?.[u.id]?.operation?.login ? [selectedManagers[u.id].operation!.login!] : []}
            onChange={onRowManagerChange(u.id, 'operation')}
            showHiddenInUI={false}
          />
        </div>
      </div>
    );
  };

  return (
    <Stack tokens={{ childrenGap: 16 }}>
      {msg && (
        <MessageBar messageBarType={msg.type} isMultiline={false} onDismiss={() => setMsg(null)}>
          {msg.text}
        </MessageBar>
      )}

      {}
      <Stack horizontal tokens={{ childrenGap: 12 }} wrap>
        <Stack style={{ minWidth: 140 }}>
          <Label>שנת הרבעון</Label>
          <TextField
            value={quarterYear}
            onChange={(_, v) => setQuarterYear((v || '').trim())}
            placeholder="לדוגמה: 2025"
            maxLength={4}
          />
        </Stack>
        <Stack style={{ minWidth: 160 }}>
          <Label>רבעון</Label>
          <Dropdown
            options={QUARTER_OPTIONS}
            selectedKey={quarterName}
            onChange={(_, opt) => opt?.key && setQuarterName(String(opt.key))}
          />
        </Stack>
      </Stack>

      <Stack tokens={{ childrenGap: 8 }}>
        <Label>רשימת עובדים מהמערכת:</Label>
        <div>כל העובדים מהרשימה נטענים אוטומטית. ניתן לערוך את כל השדות בטבלה למטה.</div>

        {selectedUsers.length > 0 && (
          <Stack tokens={{ childrenGap: 6 }}>
            <Label>סה"כ {selectedUsers.length} עובדים:</Label>

            {/* בר עליון: בחר הכל + שיוך מרוכז */}
            <Stack tokens={{ childrenGap: 12 }}>
              <Checkbox
                label="בחר / בטל בחירת כל העובדים בטבלה"
                onChange={onToggleSelectAllRows}
              />

              {/* שיוך מרוכז - סוג עובד */}
              <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="end">
                <Stack style={{ width: 180 }}>
                  <Label>סוג עובד</Label>
                  <Dropdown
                    options={WORK_TYPE_OPTIONS}
                    selectedKey={bulkWorkType}
                    onChange={(_, opt) => {
                      if (opt) setBulkWorkType(opt.key as string);
                    }}
                  />
                </Stack>

                <PrimaryButton
                  text="שיוך סוג עובד לנבחרים"
                  onClick={() => {
                    setUserWorkType(prev => {
                      const next = { ...prev };
                      selectedUsers.forEach(u => {
                        if (u.id && rowSelection[u.id]) {
                          next[u.id] = bulkWorkType;
                        }
                      });
                      return next;
                    });
                  }}
                />
              </Stack>
            </Stack>

            {/* טבלה עם גלילה */}
            <div style={{ maxHeight: 400, overflowY: 'auto', border: '1px solid #e5e7eb', borderRadius: 8, marginTop: 8 }}>
              {/* header row */}
              <div
                style={{
                  display: 'grid',
                  gridTemplateColumns: '40px 200px 180px 120px 160px 140px 150px 220px 220px 220px',
                  gap: 12,
                  padding: '8px 12px',
                  borderBottom: '1px solid #e5e7eb',
                  background: '#f5f7fa',
                  fontSize: 12,
                  fontWeight: 600,
                  alignItems: 'center'
                }}
              >
                <div />
                <div>שם עובד</div>
                <div>סוג עובד</div>
                <div>מחלקה</div>
                <div>תת-מחלקה</div>
                <div>סטטוס</div>
                <div>מנהל ישיר</div>
                <div>מנהל עקיף</div>
                <div>מנהל תפעול</div>
              </div>

              {/* rows */}
              {selectedUsers.map(u => (
                <React.Fragment key={u.id}>
                  {renderUserBadge(u)}
                </React.Fragment>
              ))}
            </div>
          </Stack>
        )}

      </Stack>

      <PrimaryButton text={busy ? 'מעדכן...' : 'שמירה ועדכון רשומות עובדים'} onClick={onSubmit} disabled={busy} />
    </Stack>
  );
};

export default EmployeeEvaluation;

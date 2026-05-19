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



const LIST_TITLE = 'employeeEvaluation';

type GroupSentPreview = { total: number; already: number; loading: boolean; };

const QUARTER_OPTIONS: IDropdownOption[] = [
  { key: 'Q1', text: 'Q1' },
  { key: 'Q2', text: 'Q2' },
  { key: 'Q3', text: 'Q3' },
  { key: 'Q4', text: 'Q4' }
];

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




//❤️
type UserMeta = {
  employeeType: string;
  managerDisplayName: string;
  managerLogin: string; // NEW: for ensureUser()
  groupNamesForSelected: string[];
  employeeNumber?: number;
};
//❤️

const EmployeeEvaluation: React.FC<IEmployeeEvaluationProps> = (props) => {
  // עובדים שנבחרו ידנית מה-PeoplePicker
  const [manualUsers, setManualUsers] = React.useState<IUser[]>([]);

  // עובדים שבאים מקבוצות: gid -> רשימת עובדים
  const [groupUsersByGroup, setGroupUsersByGroup] = React.useState<Record<string, IUser[]>>({});

  const [userWorkType, setUserWorkType] = React.useState<Record<string, string>>({});
    // בחירת עובדים בטבלה לשיוך מרוכז
  const [rowSelection, setRowSelection] = React.useState<Record<string, boolean>>({});
  const [bulkWorkType, setBulkWorkType] = React.useState<string>('רגיל');

  const { sp, graphClient, context } = props;
  const [groups, setGroups] = React.useState<IGroup[]>([]);
  const [groupOptions, setGroupOptions] = React.useState<IDropdownOption[]>([]);
  const [selectedGroupIds, setSelectedGroupIds] = React.useState<string[]>([]);
  const [selectedUsers, setSelectedUsers] = React.useState<IUser[]>([]);
  const [busy, setBusy] = React.useState(false);
  const [msg, setMsg] = React.useState<{ type: MessageBarType; text: string } | null>(null);

  // “נשלח” לפי רבעון/שנה: טוקנים
  const [sentTokens, setSentTokens] = React.useState<Set<string>>(new Set());
  const [groupPreview, setGroupPreview] = React.useState<Record<string, GroupSentPreview>>({});
  const [groupNewOnly, setGroupNewOnly] = React.useState<Record<string, boolean>>({});

  // רבעון/שנה ב-UI
  const [quarterName, setQuarterName] = React.useState<string>('Q1');
  const [quarterYear, setQuarterYear] = React.useState<string>(new Date().getFullYear().toString());

  // cache מטא למשתמש
  const userMetaCache = React.useRef<Map<string, UserMeta>>(new Map());

  const employeeNumberMapRef = React.useRef<Map<string, string> | null>(null);


  // שמות עמודות ה-User בפועל (אם קיימת התנגשות, נעבור לשמות גיבוי)
  const employeeUserFieldRef = React.useRef<string>('Employee');
  const managerUserFieldRef  = React.useRef<string>('DirectManager');

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

  React.useEffect(() => {
    recomputeSelectedUsers();
  }, [recomputeSelectedUsers]);

  React.useEffect(() => {
    (async () => {
      try {
        // רשימת המיפוי – לפי ה-GUID שנתת
        const dirList = sp.web.lists.getById('d0169395-ae9d-4173-a84a-dc3fd69d91c2');

        // חשוב: השמות כאן צריכים להתאים לשמות העמודות ברשימה!
        const items = await dirList.items
          .select('LinkTitle', 'field_6')
          .top(5000)(); // אפשר להגדיל אם צריך

        const m = new Map<string, string>();

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


  // --- קבוצות מה-Graph ---
  React.useEffect(() => {
    (async () => {
      try {
        const res = await graphClient.api('/groups?$select=id,displayName&$top=999').get();
        const raw: any[] = res?.value || [];
        const grps: IGroup[] = raw.map(g => ({ id: g.id, displayName: g.displayName }));
        grps.sort((a, b) => a.displayName.localeCompare(b.displayName, 'he'));
        setGroups(grps);
        setGroupOptions(grps.map(g => ({ key: g.id, text: g.displayName })));
      } catch (e: any) {
        setMsg({ type: MessageBarType.error, text: `טעינת קבוצות נכשלה: ${e?.message || e}` });
      }
    })();
  }, [graphClient]);

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

  // --- PeoplePicker → בחירת משתמשים ---
  const onUsersChange = (items: any[]) => {
    console.log("🫥😥🦜 items ", items);
    const mapped: IUser[] = items.map(i => ({
      id: (i.id?.toString?.() ?? i.id) as string,
      displayName: i.text ?? i.secondaryText ?? i.loginName,
      userPrincipalName: (i.secondaryText ?? i.loginName ?? i.text ?? '').toLowerCase(),
      secondaryText: i.secondaryText 
    }));
    setManualUsers(mapped);
  };

  // --- בחירת קבוצות ---
  const onGroupsChange = async (_: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => {
    if (!option) return;
    setSelectedGroupIds(prev => {
      const next = new Set(prev);
      if (option.selected) {
        next.add(option.key as string);
        setGroupNewOnly(s => ({ ...s, [option.key as string]: s[option.key as string] ?? true }));
        ensureGroupPreview(option.key as string);
        addGroupMembersToSelected(option.key as string);
      } else {
        next.delete(option.key as string);

        setGroupNewOnly(s => {
          const c = { ...s };
          delete c[option.key as string];
          return c;
        });

        // ❗ חדש: להסיר את עובדי הקבוצה הזאת ממפת הקבוצות
        setGroupUsersByGroup(prev => {
          const clone = { ...prev };
          delete clone[option.key as string];
          return clone;
        });
      }

      userMetaCache.current.clear();
      return Array.from(next);
    });
  };

  // --- פריוויו לקבוצה (already לפי רבעון/שנה ב-UI) ---
  const ensureGroupPreview = async (gid: string) => {
    setGroupPreview(prev => ({ ...prev, [gid]: prev[gid] ?? { total: 0, already: 0, loading: true } }));
    try {
      const members = await expandGroupMembers([gid]);
      const total = members.length;
      let already = 0;
      for (const u of members) {
        const k1 = makeKey(u.userPrincipalName || '', quarterName, quarterYear);
        const k2 = makeKey(u.displayName || '',       quarterName, quarterYear);
        if (sentTokens.has(k1) || sentTokens.has(k2)) already++;
      }
      setGroupPreview(prev => ({ ...prev, [gid]: { total, already, loading: false } }));
    } catch {
      setGroupPreview(prev => ({ ...prev, [gid]: { total: 0, already: 0, loading: false } }));
    }
  };

  // רענון פריוויו כשמשנים רבעון/שנה או כשהטוקנים משתנים
  React.useEffect(() => {
    if (selectedGroupIds.length === 0) return;
    selectedGroupIds.forEach(gid => ensureGroupPreview(gid));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [quarterName, quarterYear, sentTokens]);


  
  // ===== עזר: הבטחת עמודת User בשם מועדף, ואם יש התנגשויות – יצירת גיבוי =====
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
      return preferredInternalName;
    }
    // קיים אבל לא מטיפוס User – נשתמש בגיבוי
  } catch {
    // לא קיים – ננסה ליצור בשם המועדף
    try {
      await list.fields.addUser(preferredInternalName, {
        Description: description,
        SelectionMode: 0 // Single user
      });
      return preferredInternalName;
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
      return fallbackInternalName;
    }
  } catch {
    // לא קיים – ניצור
  }

  await list.fields.addUser(fallbackInternalName, {
    Description: description,
    SelectionMode: 0
  });

  return fallbackInternalName;
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

      const ensureMultilineField = async (nameOrTitle: string, opts: any) => {
        try {
          await list.fields.getByInternalNameOrTitle(nameOrTitle)();
        } catch {
          await list.fields.addMultilineText(nameOrTitle, opts);
        }
      };

      await ensureNumberField('EmployeeNameNumber');

      await ensureChoiceField('WorkType', {
        Choices: ['רגיל', 'שעתי', 'מנהל'],
        FillInChoice: false
      });

      // --- שדות טקסט/בחירה/מספר ---

      await ensureTextField('EmployeeName', {
        Description: 'שם העובד'
      });

      await ensureChoiceField('EmployeeType', {
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
      }

      await ensureTextField('Source', {
        Description: 'Selected / FromGroup:<name>',
        MaxLength: 255
      });

      await ensureTextField('GroupId', {
        MaxLength: 255
      });

      await ensureMultilineField('GroupName', {
        NumberOfLines: 6,
        RichText: false,
        AppendOnly: false
      });

      // --- הבטחת עמודות User אמיתיות לעובד ולמנהל ---
      // אם "Employee" או "DirectManager" קיימים בטיפוס שגוי — ניצור EmployeeUser / DirectManagerUser

      const employeeField = await ensureUserField(
        list,
        'Employee',
        'EmployeeUser',
        'העובד הנבחר'
      );

      const managerField = await ensureUserField(
        list,
        'DirectManager',
        'DirectManagerUser',
        'המנהל הישיר'
      );

      employeeUserFieldRef.current = employeeField;
      managerUserFieldRef.current = managerField;
  };


  // --- הרחבת חברי קבוצה ---
  const expandGroupMembers = async (groupIds: string[]): Promise<IUser[]> => {
    const users = new Map<string, IUser>();
    for (const gid of groupIds) {
      let url = `/groups/${gid}/members?$select=id,displayName,userPrincipalName&$top=999`;
      while (url) {
        const page = await graphClient.api(url).get();
        for (const m of (page?.value || [])) {
          if (m['@odata.type']?.toLowerCase?.().endsWith('user')) {
            const u: IUser = {
              id: m.id,
              displayName: m.displayName,
              userPrincipalName: (m.userPrincipalName || '').toLowerCase(),
              secondaryText : (m.secondaryText)
            };
            console.log("🦄 GROUP IUSER ", u);
            users.set(u.id, u);
          }
        }
        const next = page['@odata.nextLink'] as string | undefined;
        url = next ? next.replace('https://graph.microsoft.com/v1.0', '') : '';
      }
    }
    return Array.from(users.values());
  };

  

  const addGroupMembersToSelected = async (gid: string) => {
  try {
    const members = await expandGroupMembers([gid]);

    setGroupUsersByGroup(prev => ({
      ...prev,
      [gid]: members
    }));
  } catch (e) {
    console.warn('Failed to add group members to selectedUsers', gid, e);
  }
};



  // --- מטא־דאטה אוטומטי למשתמש ---
  const getUserMeta = async (user: IUser): Promise<UserMeta> => {
    const key = user.id || user.userPrincipalName;
    if (key && userMetaCache.current.has(key)) return userMetaCache.current.get(key)!;

    let employeeType = 'רגיל';
    let employeeNumber = '';
    console.log(employeeNumber);
     // 🔍 ניסיון להביא מספר עובד מהרשימה לפי SamAccountName
    try {
      if (employeeNumberMapRef.current) {
        // מניחים שה-UPN הוא בסגנון: sam@domain
        const upn = (user.userPrincipalName || user.secondaryText || '').toLowerCase().trim();
        if (upn) {
          const sam = upn.split('@')[0]; // "admin@ezer.com" -> "admin"
          const fromMap = employeeNumberMapRef.current.get(sam);
          if (fromMap) {
            employeeNumber = fromMap;
          }
        }
      }
    } catch (e) {
      console.warn('Failed to resolve employeeNumber from SP mapping list for user', user, e);
    }
    try {
      //const u = await graphClient.api(`/users/${encodeURIComponent(user.id || user.userPrincipalName)}`).select('employeeType,displayName,userPrincipalName').get();
      const test =  await graphClient.api(`/users/${encodeURIComponent(user.secondaryText)}`).select('*').get();
      console.log("😶‍🌫️😶‍🌫️😶‍🌫️😶‍🌫️😶‍🌫️😶‍🌫️😶‍🌫️😶‍🌫️😶‍🌫️😶‍🌫️ test ", test);
      const u = await graphClient.api(`/users/${encodeURIComponent(user.secondaryText)}`).select('employeeType,displayName,userPrincipalName').get();
      if (u?.employeeType) employeeType = u.employeeType;
      console.log("👽👽 getUserMeta u ", u);
    } catch {}

    let managerDisplayName = '';
    let managerLogin = '';
    try {
      //const m = await graphClient.api(`/users/${encodeURIComponent(user.id || user.userPrincipalName)}/manager`).select('displayName,userPrincipalName').get();
      const m = await graphClient.api(`/users/${encodeURIComponent(user.secondaryText)}/manager`).select('displayName,userPrincipalName').get();
      managerDisplayName = m?.displayName || m?.userPrincipalName || '';
      managerLogin = m?.userPrincipalName || ''; // חשוב ל-ensureUser

      console.log("👽 getUserMeta m ", m);
    } catch {}

   // --- כל הקבוצות של המשתמש (ALL group names) ---
    const groupNamesForSelected: string[] = [];
    try {
      // העדיפי UPN; אם אין – AAD ObjectId; רק בסוף id מקומי אם את באמת שומרת שם GUID של AAD.
      const userKey =
        (user.userPrincipalName && user.userPrincipalName.trim()) ||
        (user as any).secondaryText || // אם הוספת לשדה ה־IUser שלך
        user.id;                       // ודאי שזה GUID של AAD, לא מספר מ-SharePoint

      // מסננים מראש רק אובייקטים מסוג קבוצה בעזרת ה-type cast:
      // אין @odata.type ב-$select, ולכן לא נקבל 400.
      let url = `/users/${encodeURIComponent(userKey)}/transitiveMemberOf/microsoft.graph.group?$select=displayName,id&$top=999`;

      const seen = new Set<string>(); // מניעת כפילויות
      while (url) {
        const page = await graphClient.api(url).get();

        for (const g of (page?.value || [])) {
          const name = g?.displayName?.trim();
          if (name && !seen.has(name)) {
            seen.add(name);
            groupNamesForSelected.push(name);
          }
        }

        const next = page['@odata.nextLink'] as string | undefined;
        url = next ? next.replace('https://graph.microsoft.com/v1.0', '') : '';
      }

      console.log('🤖 ALL groups user is in:', groupNamesForSelected);
    } catch (e) {
      console.warn('Failed to fetch ALL group names for user:', user, e);
    }

    const meta: UserMeta = { employeeType, managerDisplayName, managerLogin, groupNamesForSelected,  employeeNumber: employeeNumber ? Number(employeeNumber) : undefined};
    if (key) userMetaCache.current.set(key, meta);
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


    const userKey = user.id || user.userPrincipalName;
    const workType = userKey ? userWorkType[userKey] : undefined;

    // הבטחת Site Users Ids לעובד ולמנהל
    const employeeLogin = user.userPrincipalName || user.displayName || '';


    const ensuredEmployee = await sp.web.ensureUser(employeeLogin);
    const employeeUserId = ensuredEmployee.Id;

    let directManagerUserId: number | null = null;
    if (meta.managerLogin) {
      try {
        const ensuredManager = await sp.web.ensureUser(meta.managerLogin);
        directManagerUserId = ensuredManager.Id;
      } catch {
        directManagerUserId = null;
      }
    }

    // שמות השדות בפועל (ייתכן שהם EmployeeUser / DirectManagerUser)
    const employeeFieldName = employeeUserFieldRef.current;   // e.g. 'Employee' or 'EmployeeUser'
    const managerFieldName  = managerUserFieldRef.current;    // e.g. 'DirectManager' or 'DirectManagerUser'

    const baseFields: any = {
      Title: upnRaw,
      Source: source,
      GroupId: groupId || null,

      EmployeeName: user.displayName || user.userPrincipalName,
      EmployeeType: workType,
      QuarterName: quarterName,
      QuarterYear: qyNum,
      Status: 'ממתין לשליחה',
      GroupName: groupNameString,
      EmployeeNameNumber: meta.employeeNumber ? Number(meta.employeeNumber) : null, 
      WorkType: workType  
    };

    // הצבה לשדות User נעשית עם סיומת Id
    baseFields[`${employeeFieldName}Id`] = employeeUserId;
    if (directManagerUserId) {
      baseFields[`${managerFieldName}Id`] = directManagerUserId;
    }

    if (existing.length === 0) {
      await list.items.add(baseFields);
    } else {
      const id = existing[0].Id;
      const updateFields: any = {
        EmployeeType: workType,
        GroupName: groupNameString || existing[0].GroupName,
        EmployeeNameNumber: meta.employeeNumber
        ? Number(meta.employeeNumber)
        : existing[0].EmployeeNameNumber, 
        WorkType: workType

      };
      updateFields[`${employeeFieldName}Id`] = employeeUserId;
      if (directManagerUserId) {
        updateFields[`${managerFieldName}Id`] = directManagerUserId;
      }
      // אפשר למחוק אם היה לך בעבר DirectManager טקסטואלי:
      // updateFields['DirectManager'] = meta.managerDisplayName || '';
      await list.items.getById(id).update(updateFields);
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
      const usersWithoutType = selectedUsers.filter(u => !userWorkType[u.id]);

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

      // 1) משתמשים נבחרים — נשלח רק אם לא נשלח כבר ברבעון/שנה הנוכחיים
      const manualById = new Map<string, IUser>();
      for (const u of manualUsers) {
        if (u?.id) manualById.set(u.id, u);
      }
      for (const u of Array.from(manualById.values())) {
        const k1 = makeKey(u.userPrincipalName || '', quarterName, quarterYear);
        const k2 = makeKey(u.displayName || '',       quarterName, quarterYear);
        if (sentTokens.has(k1) || sentTokens.has(k2)) continue;

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


        const sendOnlyNew = groupNewOnly[gid] ?? true;
        const membersToSend = sendOnlyNew
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

        await ensureGroupPreview(gid);
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
        setMsg({ type: MessageBarType.success, text: 'התהליך החל בהצלחה עבור כל העובדים שנבחרו.' });
      } else {
        const names = failures
          .slice(0, 10)
          .map(f => f.user.displayName || f.user.userPrincipalName || '(ללא שם)')
          .join(', ');
        const extra = failures.length > 10 ? ` ועוד ${failures.length - 10} נוספים` : '';
        setMsg({
          type: MessageBarType.warning,
          text: `הפעולה הושלמה חלקית: חלק מהעובדים נוספו בהצלחה, אך ${failures.length} כשלו. בעיות: ${names}${extra}. ראי לוג בקונסול לפרטים.`
        });
      }
    } catch (e: any) {
      setMsg({ type: MessageBarType.error, text: `שגיאה בשליחה: ${e?.message || e}` });
    } finally {
      setBusy(false);
    }
  };

  // ====== הדגשה ורודה ב-PeoplePicker — רק לרבעון/שנה הנוכחיים ======
  const pickerHostRef = React.useRef<HTMLDivElement | null>(null);

  React.useEffect(() => {
    const styleId = 'ao-picker-highlight-style';
    if (!document.getElementById(styleId)) {
      const style = document.createElement('style');
      style.id = styleId;
      style.textContent = `
        .ao-already-sent { background: #ffe0ef !important; border: 1px solid #ff9ec4 !important; border-radius: 6px !important; }
      `;
      document.head.appendChild(style);
    }
  }, []);

  const recolorPickerDom = React.useCallback(() => {
    if (!pickerHostRef.current) return;

    const paint = (nodeList: NodeListOf<HTMLElement>) => {
      nodeList.forEach(el => {
        const textNorm = normalize(el.textContent || '');
        const match = sentTokens.has(makeKey(textNorm, quarterName, quarterYear));
        if (match) el.classList.add('ao-already-sent');
        else el.classList.remove('ao-already-sent');
      });
    };

    const suggestionItems = pickerHostRef.current.querySelectorAll<HTMLElement>(
      `.ms-Suggestions-item, .ms-PickerPersona-container, .ms-Suggestion-item, .ms-PeoplePicker-personaContent`
    );
    paint(suggestionItems);

    const selectedItems = pickerHostRef.current.querySelectorAll<HTMLElement>(
      `.ms-PickerItem-content, .ms-PickerPersona-container, .ms-Persona-primaryText`
    );
    paint(selectedItems);
  }, [sentTokens, quarterName, quarterYear]);

  React.useEffect(() => {
    if (!pickerHostRef.current) return;
    const obs = new MutationObserver(() => recolorPickerDom());
    obs.observe(pickerHostRef.current, { childList: true, subtree: true, characterData: true });
    recolorPickerDom();
    return () => obs.disconnect();
  }, [recolorPickerDom]);

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
    const already =
      sentTokens.has(makeKey(u.userPrincipalName || '', quarterName, quarterYear)) ||
      sentTokens.has(makeKey(u.displayName || '',       quarterName, quarterYear));

    const currentWorkType = userWorkType[u.id] || '';

    return (
      <div
        style={{
          display: 'grid',
          gridTemplateColumns: '32px 1fr 1fr 140px',
          gap: 8,
          alignItems: 'center',
          padding: '4px 8px',
          borderBottom: '1px solid #e5e7eb',
          background: already ? '#ffe0ef' : 'transparent'
        }}
      >
        {/* צ׳קבוקס בחירה לשיוך מרוכז */}
        <Checkbox
          checked={!!rowSelection[u.id]}
          onChange={(_, checked) => {
            setRowSelection(prev => ({ ...prev, [u.id]: !!checked }));
          }}
        />

        {/* שם העובד */}
        <span>{u.displayName || u.userPrincipalName}</span>

        {/* מצב "כבר נשלח" + סוג נוכחי */}
        <span style={{ fontSize: 12 }}>
          {already && (
            <span
              style={{
                marginLeft: 8,
                padding: '2px 6px',
                borderRadius: 6,
                background: '#ffd6ea',
                border: '1px solid #ff9ec4'
              }}
            >
              כבר נשלח
            </span>
          )}
          {currentWorkType && (
            <span style={{ marginInlineStart: 8 }}>
              סוג עובד: <strong>{currentWorkType}</strong>
            </span>
          )}
        </span>

        {/* (אופציונלי) שיוך פרטני אם ממש תרצי להשאיר */}
        {/* אפשר למחוק את הדרופדאון הזה אם רוצים רק שיוך מרוכז */}
        <Dropdown
          styles={{ root: { minWidth: 120 } }}
          options={WORK_TYPE_OPTIONS}
          placeholder="סוג עובד"
          selectedKey={currentWorkType || undefined}
          onChange={(_, opt) => {
            if (!opt) return;
            setUserWorkType(prev => ({ ...prev, [u.id]: opt.key as string }));
          }}
        />
      </div>
    );
  };

/*
  // --- UI עזר ---
  const renderUserBadge = (u: IUser) => {
  const already =
    sentTokens.has(makeKey(u.userPrincipalName || '', quarterName, quarterYear)) ||
    sentTokens.has(makeKey(u.displayName || '',       quarterName, quarterYear));

  const currentWorkType = userWorkType[u.id] || 'רגיל';

  return (
    <div
      style={{
        display: 'inline-flex',
        gap: 8,
        alignItems: 'center',
        padding: '4px 8px',
        border: '1px solid #e5e7eb',
        borderRadius: 8,
        background: already ? '#ffe0ef' : 'transparent'
      }}
    >
      <span>{u.displayName || u.userPrincipalName}</span>
      {already && (
        <span
          style={{
            fontSize: 12,
            padding: '2px 6px',
            borderRadius: 6,
            background: '#ffd6ea',
            border: '1px solid #ff9ec4'
          }}
        >
          כבר נשלח
        </span>
      )}

      {}
      <Dropdown
        styles={{ root: { minWidth: 120 } }}
        options={WORK_TYPE_OPTIONS}
        selectedKey={currentWorkType}
        onChange={(_, opt) => {
          if (!opt) return;
          setUserWorkType(prev => ({ ...prev, [u.id]: opt.key as string }));
        }}
      />
    </div>
  );
};
*/



  const onToggleGroupNewOnly = (gid: string, checked?: boolean) => {
    setGroupNewOnly(prev => ({ ...prev, [gid]: !!checked }));
  };

  const renderGroupBadge = (gid: string) => {
  const g = groups.find(x => x.id === gid);
  const name = g?.displayName ?? gid;
  const info = groupPreview[gid];
  const isPartialSent = info && !info.loading && info.already > 0 && info.already < info.total;


  return (
    <div
      key={gid}
      style={{
        display: 'grid',
        gap: 6,
        alignItems: 'center',
        padding: '8px 10px',
        border: '1px solid ' + (isPartialSent ? '#a7f3d0' : '#e5e7eb'),
        background: isPartialSent ? '#eaffe5' : 'transparent',
        borderRadius: 8,
        gridTemplateColumns: '1fr auto'
      }}
    >
      <div style={{ display: 'inline-flex', gap: 8, alignItems: 'center' }}>
        <span>{name}</span>
        {info?.loading && (
          <span style={{ fontSize: 12, padding: '2px 6px', borderRadius: 6, background: '#fff7e6', border: '1px solid #ffe1b7' }}>
            טוען ספירה…
          </span>
        )}
        {info && !info.loading && (
          <span style={{ fontSize: 12, padding: '2px 6px', borderRadius: 6, background: '#eef2ff', border: '1px solid #c7d2fe' }}>
            כבר נשלח ל־{info.already} מתוך {info.total}
          </span>
        )}
      </div>

      <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
        <Checkbox
          label="שליחה למשתמשים שעדיין לא נבחרו"
          checked={groupNewOnly[gid] ?? true}
          onChange={(_, checked) => onToggleGroupNewOnly(gid, checked)}
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
      <Stack horizontal tokens={{ childrenGap: 12 }}>
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
        <Label>בחירת עובדים פעילים:</Label>
        <div ref={pickerHostRef}>
          <PeoplePicker
            context={peoplePickerContext}
            personSelectionLimit={50}
            principalTypes={[PrincipalType.User]}
            ensureUser={true}
            onChange={onUsersChange}
            showHiddenInUI={false}
          />
        </div>

        {selectedUsers.length > 0 && (
          <Stack tokens={{ childrenGap: 6 }}>
            <Label>נבחרו עובדים:</Label>

            {/* בר עליון: בחר הכל + סוג עובד מרוכז + כפתור שיוך */}
            <Stack horizontal tokens={{ childrenGap: 12 }} verticalAlign="center">
              <Checkbox
                label="בחר / בטל בחירת כל העובדים בטבלה"
                onChange={onToggleSelectAllRows}
              />

              <Dropdown
                options={WORK_TYPE_OPTIONS}
                selectedKey={bulkWorkType}
                styles={{ root: { width: 180 } }}
                onChange={(_, opt) => {
                  if (opt) setBulkWorkType(opt.key as string);
                }}
              />

              <PrimaryButton
                text="שיוך לסוג עובד הנבחר"
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

            {/* טבלה עם גלילה */}
            <div style={{ maxHeight: 300, overflowY: 'auto', border: '1px solid #e5e7eb', borderRadius: 8, marginTop: 8 }}>
              {selectedUsers.map(u => (
                <React.Fragment key={u.id}>{renderUserBadge(u)}</React.Fragment>
              ))}
            </div>
          </Stack>
        )}

      </Stack>

      <Stack tokens={{ childrenGap: 8 }}>
        <Label>בחירת קבוצות פעילות:</Label>
        <Dropdown placeholder="בחרי קבוצות" multiSelect options={groupOptions} onChange={onGroupsChange} />
        {selectedGroupIds.length > 0 && (
          <Stack tokens={{ childrenGap: 6 }}>
            <Label>נבחרו קבוצות:</Label>
            <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
              {selectedGroupIds.map(renderGroupBadge)}
            </div>
          </Stack>
        )}
      </Stack>

      <PrimaryButton text={busy ? 'שולח...' : 'התחלת תהליך הערכת עובדים'} onClick={onSubmit} disabled={busy} />
    </Stack>
  );
};

export default EmployeeEvaluation;

/*
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



const LIST_TITLE = 'employeeEvaluation';

type GroupSentPreview = { total: number; already: number; loading: boolean; };

const QUARTER_OPTIONS: IDropdownOption[] = [
  { key: 'Q1', text: 'Q1' },
  { key: 'Q2', text: 'Q2' },
  { key: 'Q3', text: 'Q3' },
  { key: 'Q4', text: 'Q4' }
];

const STATUS_CHOICES = [
  'ממתין לשליחה',
  'נשלח',
  'מולא ע"י העובד',
  'מולא על יד המנהל',
  'אושר',
  'נדחה',
  'נשלח לתיקון'
];

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




//❤️
type UserMeta = {
  employeeType: string;
  managerDisplayName: string;
  managerLogin: string; // NEW: for ensureUser()
  groupNamesForSelected: string[];
  employeeNumber?: number;
};
//❤️

const EmployeeEvaluation: React.FC<IEmployeeEvaluationProps> = (props) => {
  const { sp, graphClient, context } = props;
  const [groups, setGroups] = React.useState<IGroup[]>([]);
  const [groupOptions, setGroupOptions] = React.useState<IDropdownOption[]>([]);
  const [selectedGroupIds, setSelectedGroupIds] = React.useState<string[]>([]);
  const [selectedUsers, setSelectedUsers] = React.useState<IUser[]>([]);
  const [busy, setBusy] = React.useState(false);
  const [msg, setMsg] = React.useState<{ type: MessageBarType; text: string } | null>(null);

  // “נשלח” לפי רבעון/שנה: טוקנים
  const [sentTokens, setSentTokens] = React.useState<Set<string>>(new Set());
  const [groupPreview, setGroupPreview] = React.useState<Record<string, GroupSentPreview>>({});
  const [groupNewOnly, setGroupNewOnly] = React.useState<Record<string, boolean>>({});

  // רבעון/שנה ב-UI
  const [quarterName, setQuarterName] = React.useState<string>('Q1');
  const [quarterYear, setQuarterYear] = React.useState<string>(new Date().getFullYear().toString());

  // cache מטא למשתמש
  const userMetaCache = React.useRef<Map<string, UserMeta>>(new Map());

  const employeeNumberMapRef = React.useRef<Map<string, string> | null>(null);


  // שמות עמודות ה-User בפועל (אם קיימת התנגשות, נעבור לשמות גיבוי)
  const employeeUserFieldRef = React.useRef<string>('Employee');
  const managerUserFieldRef  = React.useRef<string>('DirectManager');

  // PeoplePicker context
  const peoplePickerContext: IPeoplePickerContext = {
    absoluteUrl: context.pageContext.web.absoluteUrl,
    spHttpClient: context.spHttpClient,
    msGraphClientFactory: context.msGraphClientFactory
  };

  React.useEffect(() => {
    (async () => {
      try {
        // רשימת המיפוי – לפי ה-GUID שנתת
        const dirList = sp.web.lists.getById('d0169395-ae9d-4173-a84a-dc3fd69d91c2');

        // חשוב: השמות כאן צריכים להתאים לשמות העמודות ברשימה!
        const items = await dirList.items
          .select('LinkTitle', 'field_6')
          .top(5000)(); // אפשר להגדיל אם צריך

        const m = new Map<string, string>();

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


  // --- קבוצות מה-Graph ---
  React.useEffect(() => {
    (async () => {
      try {
        const res = await graphClient.api('/groups?$select=id,displayName&$top=999').get();
        const raw: any[] = res?.value || [];
        const grps: IGroup[] = raw.map(g => ({ id: g.id, displayName: g.displayName }));
        grps.sort((a, b) => a.displayName.localeCompare(b.displayName, 'he'));
        setGroups(grps);
        setGroupOptions(grps.map(g => ({ key: g.id, text: g.displayName })));
      } catch (e: any) {
        setMsg({ type: MessageBarType.error, text: `טעינת קבוצות נכשלה: ${e?.message || e}` });
      }
    })();
  }, [graphClient]);

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

  // --- PeoplePicker → בחירת משתמשים ---
  const onUsersChange = (items: any[]) => {
    console.log("🫥😥🦜 items ", items);
    const mapped: IUser[] = items.map(i => ({
      id: (i.id?.toString?.() ?? i.id) as string,
      displayName: i.text ?? i.secondaryText ?? i.loginName,
      userPrincipalName: (i.secondaryText ?? i.loginName ?? i.text ?? '').toLowerCase(),
      secondaryText: i.secondaryText 
    }));
    setSelectedUsers(mapped);
  };

  // --- בחירת קבוצות ---
  const onGroupsChange = async (_: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => {
    if (!option) return;
    setSelectedGroupIds(prev => {
      const next = new Set(prev);
      if (option.selected) {
        next.add(option.key as string);
        setGroupNewOnly(s => ({ ...s, [option.key as string]: s[option.key as string] ?? true }));
        ensureGroupPreview(option.key as string);
      } else {
        next.delete(option.key as string);
        setGroupNewOnly(s => {
          const c = { ...s };
          delete c[option.key as string];
          return c;
        });
      }
      userMetaCache.current.clear();
      return Array.from(next);
    });
  };

  // --- פריוויו לקבוצה (already לפי רבעון/שנה ב-UI) ---
  const ensureGroupPreview = async (gid: string) => {
    setGroupPreview(prev => ({ ...prev, [gid]: prev[gid] ?? { total: 0, already: 0, loading: true } }));
    try {
      const members = await expandGroupMembers([gid]);
      const total = members.length;
      let already = 0;
      for (const u of members) {
        const k1 = makeKey(u.userPrincipalName || '', quarterName, quarterYear);
        const k2 = makeKey(u.displayName || '',       quarterName, quarterYear);
        if (sentTokens.has(k1) || sentTokens.has(k2)) already++;
      }
      setGroupPreview(prev => ({ ...prev, [gid]: { total, already, loading: false } }));
    } catch {
      setGroupPreview(prev => ({ ...prev, [gid]: { total: 0, already: 0, loading: false } }));
    }
  };

  // רענון פריוויו כשמשנים רבעון/שנה או כשהטוקנים משתנים
  React.useEffect(() => {
    if (selectedGroupIds.length === 0) return;
    selectedGroupIds.forEach(gid => ensureGroupPreview(gid));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [quarterName, quarterYear, sentTokens]);


  
  // ===== עזר: הבטחת עמודת User בשם מועדף, ואם יש התנגשויות – יצירת גיבוי =====
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
      return preferredInternalName;
    }
    // קיים אבל לא מטיפוס User – נשתמש בגיבוי
  } catch {
    // לא קיים – ננסה ליצור בשם המועדף
    try {
      await list.fields.addUser(preferredInternalName, {
        Description: description,
        SelectionMode: 0 // Single user
      });
      return preferredInternalName;
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
      return fallbackInternalName;
    }
  } catch {
    // לא קיים – ניצור
  }

  await list.fields.addUser(fallbackInternalName, {
    Description: description,
    SelectionMode: 0
  });

  return fallbackInternalName;
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

      const ensureMultilineField = async (nameOrTitle: string, opts: any) => {
        try {
          await list.fields.getByInternalNameOrTitle(nameOrTitle)();
        } catch {
          await list.fields.addMultilineText(nameOrTitle, opts);
        }
      };

      await ensureNumberField('EmployeeNameNumber');

      // --- שדות טקסט/בחירה/מספר ---

      await ensureTextField('EmployeeName', {
        Description: 'שם העובד'
      });

      await ensureChoiceField('EmployeeType', {
        Choices: ['עובד', 'קבלן', 'סטודנט', 'אחר'],
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
      }

      await ensureTextField('Source', {
        Description: 'Selected / FromGroup:<name>',
        MaxLength: 255
      });

      await ensureTextField('GroupId', {
        MaxLength: 255
      });

      await ensureMultilineField('GroupName', {
        NumberOfLines: 6,
        RichText: false,
        AppendOnly: false
      });

      // --- הבטחת עמודות User אמיתיות לעובד ולמנהל ---
      // אם "Employee" או "DirectManager" קיימים בטיפוס שגוי — ניצור EmployeeUser / DirectManagerUser

      const employeeField = await ensureUserField(
        list,
        'Employee',
        'EmployeeUser',
        'העובד הנבחר'
      );

      const managerField = await ensureUserField(
        list,
        'DirectManager',
        'DirectManagerUser',
        'המנהל הישיר'
      );

      employeeUserFieldRef.current = employeeField;
      managerUserFieldRef.current = managerField;
  };


  // --- הרחבת חברי קבוצה ---
  const expandGroupMembers = async (groupIds: string[]): Promise<IUser[]> => {
    const users = new Map<string, IUser>();
    for (const gid of groupIds) {
      let url = `/groups/${gid}/members?$select=id,displayName,userPrincipalName&$top=999`;
      while (url) {
        const page = await graphClient.api(url).get();
        for (const m of (page?.value || [])) {
          if (m['@odata.type']?.toLowerCase?.().endsWith('user')) {
            const u: IUser = {
              id: m.id,
              displayName: m.displayName,
              userPrincipalName: (m.userPrincipalName || '').toLowerCase(),
              secondaryText : (m.secondaryText)
            };
            console.log("🦄 GROUP IUSER ", u);
            users.set(u.id, u);
          }
        }
        const next = page['@odata.nextLink'] as string | undefined;
        url = next ? next.replace('https://graph.microsoft.com/v1.0', '') : '';
      }
    }
    return Array.from(users.values());
  };

  // --- מטא־דאטה אוטומטי למשתמש ---
  const getUserMeta = async (user: IUser): Promise<UserMeta> => {
    const key = user.id || user.userPrincipalName;
    if (key && userMetaCache.current.has(key)) return userMetaCache.current.get(key)!;

    let employeeType = 'אחר';
    let employeeNumber = '';
    console.log(employeeNumber);
     // 🔍 ניסיון להביא מספר עובד מהרשימה לפי SamAccountName
    try {
      if (employeeNumberMapRef.current) {
        // מניחים שה-UPN הוא בסגנון: sam@domain
        const upn = (user.userPrincipalName || user.secondaryText || '').toLowerCase().trim();
        if (upn) {
          const sam = upn.split('@')[0]; // "admin@ezer.com" -> "admin"
          const fromMap = employeeNumberMapRef.current.get(sam);
          if (fromMap) {
            employeeNumber = fromMap;
          }
        }
      }
    } catch (e) {
      console.warn('Failed to resolve employeeNumber from SP mapping list for user', user, e);
    }
    try {
      //const u = await graphClient.api(`/users/${encodeURIComponent(user.id || user.userPrincipalName)}`).select('employeeType,displayName,userPrincipalName').get();
      const test =  await graphClient.api(`/users/${encodeURIComponent(user.secondaryText)}`).select('*').get();
      console.log("😶‍🌫️😶‍🌫️😶‍🌫️😶‍🌫️😶‍🌫️😶‍🌫️😶‍🌫️😶‍🌫️😶‍🌫️😶‍🌫️ test ", test);
      const u = await graphClient.api(`/users/${encodeURIComponent(user.secondaryText)}`).select('employeeType,displayName,userPrincipalName').get();
      if (u?.employeeType) employeeType = u.employeeType;
      console.log("👽👽 getUserMeta u ", u);
    } catch {}

    let managerDisplayName = '';
    let managerLogin = '';
    try {
      //const m = await graphClient.api(`/users/${encodeURIComponent(user.id || user.userPrincipalName)}/manager`).select('displayName,userPrincipalName').get();
      const m = await graphClient.api(`/users/${encodeURIComponent(user.secondaryText)}/manager`).select('displayName,userPrincipalName').get();
      managerDisplayName = m?.displayName || m?.userPrincipalName || '';
      managerLogin = m?.userPrincipalName || ''; // חשוב ל-ensureUser

      console.log("👽 getUserMeta m ", m);
    } catch {}

   // --- כל הקבוצות של המשתמש (ALL group names) ---
    const groupNamesForSelected: string[] = [];
    try {
      // העדיפי UPN; אם אין – AAD ObjectId; רק בסוף id מקומי אם את באמת שומרת שם GUID של AAD.
      const userKey =
        (user.userPrincipalName && user.userPrincipalName.trim()) ||
        (user as any).secondaryText || // אם הוספת לשדה ה־IUser שלך
        user.id;                       // ודאי שזה GUID של AAD, לא מספר מ-SharePoint

      // מסננים מראש רק אובייקטים מסוג קבוצה בעזרת ה-type cast:
      // אין @odata.type ב-$select, ולכן לא נקבל 400.
      let url = `/users/${encodeURIComponent(userKey)}/transitiveMemberOf/microsoft.graph.group?$select=displayName,id&$top=999`;

      const seen = new Set<string>(); // מניעת כפילויות
      while (url) {
        const page = await graphClient.api(url).get();

        for (const g of (page?.value || [])) {
          const name = g?.displayName?.trim();
          if (name && !seen.has(name)) {
            seen.add(name);
            groupNamesForSelected.push(name);
          }
        }

        const next = page['@odata.nextLink'] as string | undefined;
        url = next ? next.replace('https://graph.microsoft.com/v1.0', '') : '';
      }

      console.log('🤖 ALL groups user is in:', groupNamesForSelected);
    } catch (e) {
      console.warn('Failed to fetch ALL group names for user:', user, e);
    }

    const meta: UserMeta = { employeeType, managerDisplayName, managerLogin, groupNamesForSelected,  employeeNumber: employeeNumber ? Number(employeeNumber) : undefined};
    if (key) userMetaCache.current.set(key, meta);
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

    // הבטחת Site Users Ids לעובד ולמנהל
    const employeeLogin = user.userPrincipalName || user.displayName || '';
    const ensuredEmployee = await sp.web.ensureUser(employeeLogin);
    const employeeUserId = ensuredEmployee.Id;

    let directManagerUserId: number | null = null;
    if (meta.managerLogin) {
      try {
        const ensuredManager = await sp.web.ensureUser(meta.managerLogin);
        directManagerUserId = ensuredManager.Id;
      } catch {
        directManagerUserId = null;
      }
    }

    // שמות השדות בפועל (ייתכן שהם EmployeeUser / DirectManagerUser)
    const employeeFieldName = employeeUserFieldRef.current;   // e.g. 'Employee' or 'EmployeeUser'
    const managerFieldName  = managerUserFieldRef.current;    // e.g. 'DirectManager' or 'DirectManagerUser'

    const baseFields: any = {
      Title: upnRaw,
      Source: source,
      GroupId: groupId || null,

      EmployeeName: user.displayName || user.userPrincipalName,
      EmployeeType: meta.employeeType || 'אחר',
      QuarterName: quarterName,
      QuarterYear: qyNum,
      Status: 'ממתין לשליחה',
      GroupName: groupNameString,
      EmployeeNameNumber: meta.employeeNumber ? Number(meta.employeeNumber) : null
    };

    // הצבה לשדות User נעשית עם סיומת Id
    baseFields[`${employeeFieldName}Id`] = employeeUserId;
    if (directManagerUserId) {
      baseFields[`${managerFieldName}Id`] = directManagerUserId;
    }

    if (existing.length === 0) {
      await list.items.add(baseFields);
    } else {
      const id = existing[0].Id;
      const updateFields: any = {
        EmployeeType: meta.employeeType || 'אחר',
        GroupName: groupNameString || existing[0].GroupName,
        EmployeeNameNumber: meta.employeeNumber
        ? Number(meta.employeeNumber)
        : existing[0].EmployeeNameNumber
      };
      updateFields[`${employeeFieldName}Id`] = employeeUserId;
      if (directManagerUserId) {
        updateFields[`${managerFieldName}Id`] = directManagerUserId;
      }
      // אפשר למחוק אם היה לך בעבר DirectManager טקסטואלי:
      // updateFields['DirectManager'] = meta.managerDisplayName || '';
      await list.items.getById(id).update(updateFields);
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

      await ensureList();

      const actuallySent: IUser[] = [];
      const failures: { user: IUser; error: any }[] = [];

      // 1) משתמשים נבחרים — נשלח רק אם לא נשלח כבר ברבעון/שנה הנוכחיים
      const manualById = new Map<string, IUser>();
      for (const u of selectedUsers) {
        if (u?.id) manualById.set(u.id, u);
      }
      for (const u of Array.from(manualById.values())) {
        const k1 = makeKey(u.userPrincipalName || '', quarterName, quarterYear);
        const k2 = makeKey(u.displayName || '',       quarterName, quarterYear);
        if (sentTokens.has(k1) || sentTokens.has(k2)) continue;

        const r = await tryAddWorker(u, 'Selected', undefined);
        if (r.ok) actuallySent.push(u);
        else failures.push({ user: u, error: r.error });
      }

      // 2) קבוצות (מסונן לפי sentTokens לרבעון/שנה הנוכחיים)
      for (const gid of selectedGroupIds) {
        const g = groups.find(x => x.id === gid);
        const gName = g?.displayName ?? gid;
        let members: IUser[] = [];
        try {
          members = await expandGroupMembers([gid]);
        } catch (e) {
          console.warn('expandGroupMembers failed', gid, e);
          continue;
        }

        const sendOnlyNew = groupNewOnly[gid] ?? true;
        const membersToSend = sendOnlyNew
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

        await ensureGroupPreview(gid);
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
        setMsg({ type: MessageBarType.success, text: 'התהליך החל בהצלחה עבור כל העובדים שנבחרו.' });
      } else {
        const names = failures
          .slice(0, 10)
          .map(f => f.user.displayName || f.user.userPrincipalName || '(ללא שם)')
          .join(', ');
        const extra = failures.length > 10 ? ` ועוד ${failures.length - 10} נוספים` : '';
        setMsg({
          type: MessageBarType.warning,
          text: `הפעולה הושלמה חלקית: חלק מהעובדים נוספו בהצלחה, אך ${failures.length} כשלו. בעיות: ${names}${extra}. ראי לוג בקונסול לפרטים.`
        });
      }
    } catch (e: any) {
      setMsg({ type: MessageBarType.error, text: `שגיאה בשליחה: ${e?.message || e}` });
    } finally {
      setBusy(false);
    }
  };

  // ====== הדגשה ורודה ב-PeoplePicker — רק לרבעון/שנה הנוכחיים ======
  const pickerHostRef = React.useRef<HTMLDivElement | null>(null);

  React.useEffect(() => {
    const styleId = 'ao-picker-highlight-style';
    if (!document.getElementById(styleId)) {
      const style = document.createElement('style');
      style.id = styleId;
      style.textContent = `
        .ao-already-sent { background: #ffe0ef !important; border: 1px solid #ff9ec4 !important; border-radius: 6px !important; }
      `;
      document.head.appendChild(style);
    }
  }, []);

  const recolorPickerDom = React.useCallback(() => {
    if (!pickerHostRef.current) return;

    const paint = (nodeList: NodeListOf<HTMLElement>) => {
      nodeList.forEach(el => {
        const textNorm = normalize(el.textContent || '');
        const match = sentTokens.has(makeKey(textNorm, quarterName, quarterYear));
        if (match) el.classList.add('ao-already-sent');
        else el.classList.remove('ao-already-sent');
      });
    };

    const suggestionItems = pickerHostRef.current.querySelectorAll<HTMLElement>(
      `.ms-Suggestions-item, .ms-PickerPersona-container, .ms-Suggestion-item, .ms-PeoplePicker-personaContent`
    );
    paint(suggestionItems);

    const selectedItems = pickerHostRef.current.querySelectorAll<HTMLElement>(
      `.ms-PickerItem-content, .ms-PickerPersona-container, .ms-Persona-primaryText`
    );
    paint(selectedItems);
  }, [sentTokens, quarterName, quarterYear]);

  React.useEffect(() => {
    if (!pickerHostRef.current) return;
    const obs = new MutationObserver(() => recolorPickerDom());
    obs.observe(pickerHostRef.current, { childList: true, subtree: true, characterData: true });
    recolorPickerDom();
    return () => obs.disconnect();
  }, [recolorPickerDom]);

  // --- UI עזר ---
  const renderUserBadge = (u: IUser) => {
    const already =
      sentTokens.has(makeKey(u.userPrincipalName || '', quarterName, quarterYear)) ||
      sentTokens.has(makeKey(u.displayName || '',       quarterName, quarterYear));
    return (
      <div style={{ display: 'inline-flex', gap: 8, alignItems: 'center', padding: '4px 8px', border: '1px solid #e5e7eb', borderRadius: 8, background: already ? '#ffe0ef' : 'transparent' }}>
        <span>{u.displayName || u.userPrincipalName}</span>
        {already && <span style={{ fontSize: 12, padding: '2px 6px', borderRadius: 6, background: '#ffd6ea', border: '1px solid #ff9ec4' }}>כבר נשלח</span>}
      </div>
    );
  };

  const onToggleGroupNewOnly = (gid: string, checked?: boolean) => {
    setGroupNewOnly(prev => ({ ...prev, [gid]: !!checked }));
  };

  const renderGroupBadge = (gid: string) => {
    const g = groups.find(x => x.id === gid);
    const name = g?.displayName ?? gid;
    const info = groupPreview[gid];
    const isPartialSent = info && !info.loading && info.already > 0 && info.already < info.total;

    return (
      <div
        key={gid}
        style={{
          display: 'grid',
          gap: 6,
          alignItems: 'center',
          padding: '8px 10px',
          border: '1px solid ' + (isPartialSent ? '#a7f3d0' : '#e5e7eb'),
          background: isPartialSent ? '#eaffe5' : 'transparent',
          borderRadius: 8,
          gridTemplateColumns: '1fr auto'
        }}
      >
        <div style={{ display: 'inline-flex', gap: 8, alignItems: 'center' }}>
          <span>{name}</span>
          {info?.loading && (
            <span style={{ fontSize: 12, padding: '2px 6px', borderRadius: 6, background: '#fff7e6', border: '1px solid #ffe1b7' }}>
              טוען ספירה…
            </span>
          )}
          {info && !info.loading && (
            <span style={{ fontSize: 12, padding: '2px 6px', borderRadius: 6, background: '#eef2ff', border: '1px solid #c7d2fe' }}>
              כבר נשלח ל־{info.already} מתוך {info.total}
            </span>
          )}
        </div>

        <Checkbox
          label="שליחה למשתמשים שעדיין לא נבחרו"
          checked={groupNewOnly[gid] ?? true}
          onChange={(_, checked) => onToggleGroupNewOnly(gid, checked)}
        />
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
      <Stack horizontal tokens={{ childrenGap: 12 }}>
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
        <Label>בחירת עובדים פעילים:</Label>
        <div ref={pickerHostRef}>
          <PeoplePicker
            context={peoplePickerContext}
            personSelectionLimit={50}
            principalTypes={[PrincipalType.User]}
            ensureUser={true}
            onChange={onUsersChange}
            showHiddenInUI={false}
          />
        </div>

        {selectedUsers.length > 0 && (
          <Stack tokens={{ childrenGap: 6 }}>
            <Label>נבחרו עובדים:</Label>
            <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8 }}>
              {selectedUsers.map(u => <React.Fragment key={u.id}>{renderUserBadge(u)}</React.Fragment>)}
            </div>
          </Stack>
        )}
      </Stack>

      <Stack tokens={{ childrenGap: 8 }}>
        <Label>בחירת קבוצות פעילות:</Label>
        <Dropdown placeholder="בחרי קבוצות" multiSelect options={groupOptions} onChange={onGroupsChange} />
        {selectedGroupIds.length > 0 && (
          <Stack tokens={{ childrenGap: 6 }}>
            <Label>נבחרו קבוצות:</Label>
            <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
              {selectedGroupIds.map(renderGroupBadge)}
            </div>
          </Stack>
        )}
      </Stack>

      <PrimaryButton text={busy ? 'שולח...' : 'התחלת תהליך הערכת עובדים'} onClick={onSubmit} disabled={busy} />
    </Stack>
  );
};

export default EmployeeEvaluation;

*/
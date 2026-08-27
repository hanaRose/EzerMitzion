// export interface IQuickLinksProps {
//   description: string;
//   isDarkTheme: boolean;
//   environmentMessage: string;
//   hasTeamsContext: boolean;
//   userDisplayName: string;
// }

import { ListService } from '../services/ListService';

export interface IQuickLinksProps {
  listService: ListService;
}

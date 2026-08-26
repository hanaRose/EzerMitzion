// export interface ISubjectCardsProps {
//   description: string;
//   isDarkTheme: boolean;
//   environmentMessage: string;
//   hasTeamsContext: boolean;
//   userDisplayName: string;
// }

import { ListService } from '../services/ListService';

export interface ISubjectCardsProps {
  listService: ListService;
}
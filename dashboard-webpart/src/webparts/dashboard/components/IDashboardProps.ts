 import { ListService } from '../services/ListService';
import { HappyMomentsService } from '../services/HappyMomentsService';

export interface IDashboardProps {
  listService: ListService;
  happyMomentsService: HappyMomentsService;
}